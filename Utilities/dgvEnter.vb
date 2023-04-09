'------------------------------------------------
'Name: Class dgvEnter.vb.
'Function: DataGridView with overridden default Enter key behaviour changing it to move right and Drag and Drop.
'Shows a block on the row header if the data is dirty and should be updated.
'Copyright Robin Baines 2010. All rights reserved.
'------------------------------------------------
Imports System.Windows.Forms
Imports System.Windows.Forms.DataGridViewAdvancedBorderStyle
Imports System.Drawing
Public Class dgvEnter
    Inherits DataGridView

    '20121005 store the last key for dialogs.
    Dim _lastKey As Keys
    Public ReadOnly Property LastKey() As Keys
        Get
            Return _lastKey
        End Get
    End Property

    Friend WithEvents __ta As Object = Nothing
    Public Property ta() As Object
        Get
            Return __ta
        End Get
        Set(ByVal value As Object)
            __ta = value
        End Set
    End Property

    'When a Parent dg needs to fill child grids and when that takes some time the filtering slows things down.
    'Everytime a character is entered the children are filled.
    'To prevent this a timer is started on every RowEnter event and the call back to load the children (frmStandard.RowEnterTimerEvent) 
    'is called when the timer fires.
    'Start the timer functionality by setting RowEnterInterval to a time in msecs.
    'Load the children from frmStandard by overriding frmStandard.RowEnterTimerEvent.
    'If you have 2 dgs using this check which one is calling back with 
    'Public Overrides Sub RowEnterTimerEvent(dg As dgvEnter)
    'If dg.Equals(dg_drugid__receipt) Then
    Private WithEvents RowEnterTimer As New System.Windows.Forms.Timer()

    Dim _rowindex As Integer = -1
    Public Property RowIndex As Integer
        Get
            Return _rowindex
        End Get
        Set(value As Integer)
            _rowindex = value
        End Set
    End Property

    Dim _rowenterinterval As Integer
    Public Property RowEnterInterval As Integer
        Get
            Return _rowenterinterval
        End Get
        Set(value As Integer)
            _rowenterinterval = value
        End Set
    End Property

    Dim frm As frmStandard = Nothing
    Public ReadOnly Property frmParent As frmStandard
        Get
            If frm Is Nothing Then

                frm = TryCast(Me.FindForm(), frmStandard)

                'If Not frm Is Nothing Then
                '    frm.RowEnterTimerEvent(Me)
                'End If
            End If
            Return frm
        End Get
    End Property


    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal _bRO As Boolean)
        MyBase.New()
        blnRO = _bRO
        DoubleBuffered = True
    End Sub

    Private Sub Me_RowEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Me.RowEnter
        RowIndex = e.RowIndex
        If RowEnterInterval > 0 Then
            RowEnterTimer.Stop()
            RowEnterTimer.Interval = RowEnterInterval
            RowEnterTimer.Start() 'if already running only resets after a stop
        End If
    End Sub

    Protected Overridable Sub TimerEventProcessor(ByVal myObject As Object, ByVal myEventArgs As EventArgs) Handles RowEnterTimer.Tick
        RowEnterTimer.Stop()
        If Not frmParent Is Nothing Then
            frmParent.RowEnterTimerEvent(Me)
        End If
    End Sub

    Protected Overridable Sub dg_CellValidating(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles Me.CellValidating
        Dim strT As String = Me.Rows(e.RowIndex).Cells(e.ColumnIndex).FormattedValue.ToString()
        Dim strT2 = Me.Rows(e.RowIndex).Cells(e.ColumnIndex).EditedFormattedValue.ToString()
        If strT <> strT2 Then
            _blnDirty = True
        End If
    End Sub


    'Private Sub dg_CellValidated(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Me.CellValidated
    '    Dim strT As String = Me.Rows(e.RowIndex).Cells(e.ColumnIndex).FormattedValue
    '    Dim strT2 = Me.Rows(e.RowIndex).Cells(e.ColumnIndex).EditedFormattedValue
    '    If strT <> strT2 Then
    '        ' _blnDirty = True
    '    End If
    'End Sub

    '20120307 Added this but did not use in the end.
    Public Function IsVerticalScrollBarVisible() As Boolean
        Return Me.VerticalScrollBar.Visible
    End Function

    Protected Overrides Function ProcessDialogKey(ByVal keyData As Keys) As Boolean

        If keyData = Keys.Enter Then
            Dim CurrentIndex As Integer = Me.CurrentCell.ColumnIndex
            Dim blnRet As Boolean = Me.ProcessRightKey(keyData)

            'When editing it is important that the cell does move to another one so if this is the last cell
            'send it to the left.
            '20100917 can fail if user enters a new value while filter is set and there is only one record.
            Try
                '20111126 RPB changed this because calling ProcessRightKey and then ProcessLeftKey calls CellValidating
                'twice.
                'If CurrentIndex = Me.CurrentCell.ColumnIndex Then
                If blnRet = False Then
                    Me.ProcessLeftKey(keyData)
                    If CurrentIndex = Me.CurrentCell.ColumnIndex Then

                        'This occurs if there is only one field.
                        If Me.EndEdit() = True Then
                            _blnDirty = True
                        End If
                    End If
                End If

            Catch ex As Exception
            End Try
            Return blnRet
        End If

        Return MyBase.ProcessDialogKey(keyData)
    End Function

    Protected Overrides Function ProcessDataGridViewKey(ByVal e As KeyEventArgs) As Boolean
        Try
            '20121005 store the last key for dialogs.
            _lastKey = e.KeyCode
            If e.KeyCode = Keys.Enter Then
                Dim blnRet = Me.ProcessRightKey(e.KeyData)
                Return blnRet
            End If
            If e.KeyCode = Keys.Escape Then
                _blnDirty = False
            End If

            Return MyBase.ProcessDataGridViewKey(e)
        Catch ex As Exception

        End Try
        'If try fails return false.
        Return False
    End Function

    'Private Sub dg_EditingControlShowing(ByVal sender As Object, ByVal e As DataGridViewEditingControlShowingEventArgs) Handles Me.EditingControlShowing
    '    ' e.CellStyle.BackColor = Color.Aquamarine
    'End Sub
    Private _blnDirty As Boolean
    Public Property blnDirty() As Boolean
        Get
            Return _blnDirty
        End Get
        Set(ByVal value As Boolean)
            _blnDirty = value
        End Set
    End Property

    'Called from the form if the user has clicked the update (tsb) button.
    Public Overridable Sub UpdateData()
        If Me.EndEdit() = True Then

            'If Endedit did something.
            Dim bs As BindingSource = CType(Me.DataSource, BindingSource)
            bs.EndEdit()
            Dim sD As System.Data.DataRowView
            sD = bs.Current

            'Used the table adapter stored in the dgvEnter object to update the Row.
            If Not Me.ta Is Nothing Then
                Try
                    Me.ta.update(sD.Row)
                Catch ex As Exception
                End Try
            End If
            _blnDirty = False
            Me.Refresh()
        End If
    End Sub

    'Called from the form if the user has clicked the update (tsb) button.
    Public Sub UpdateAllData()

        Dim bs As BindingSource = CType(Me.DataSource, BindingSource)
        bs.EndEdit()

        'Used the table adapter stored in the dgvEnter object to update the Row.
        If Not Me.ta Is Nothing Then
            Me.ta.update(bs.DataSource)
        End If
        _blnDirty = False
        Me.Refresh()

    End Sub

    'Show a rect in the row header column if the data has been altered.
    Protected Overrides Sub OnRowPostPaint(ByVal e As DataGridViewRowPostPaintEventArgs)
        If Not Me.CurrentRow Is Nothing Then
            If e.RowIndex = Me.CurrentRow.Index And _blnDirty = True Then

                Dim strRowNumber As String = (e.RowIndex + 1).ToString()
                Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)

                If Me.RowHeadersWidth > 20 Then

                    Dim rect As New RectangleF(e.RowBounds.Location.X + Me.RowHeadersWidth - 20, e.RowBounds.Location.Y + 5, 10, e.RowBounds.Height - 10)
                    e.Graphics.FillRectangle(Brushes.Gray, rect)
                End If
            End If
        End If
        MyBase.OnRowPostPaint(e)
    End Sub


#Region "switchOffsorting"
    Public Sub Switch_off_sorting()
        For Each col As DataGridViewColumn In Columns
            col.SortMode = DataGridViewColumnSortMode.NotSortable
        Next
    End Sub
#End Region

#Region "drag and drop"

    Protected SourceCell As DataGridView.HitTestInfo = Nothing

    Private _blnMeIsSource As Boolean = False
    Public Property blnMeIsSource() As Boolean
        Get
            Return _blnMeIsSource
        End Get
        Set(ByVal value As Boolean)
            _blnMeIsSource = value
        End Set
    End Property

    Private _blnMove As Boolean = True
    Public Property blnMove() As Boolean
        Get
            Return _blnMove
        End Get
        Set(ByVal value As Boolean)
            _blnMove = value
        End Set
    End Property

    'Is only valid if _blnMeIsSource = true
    Public ReadOnly Property SourceRow() As Integer
        Get
            Return SourceCell.RowIndex
        End Get
    End Property
    'Is only valid if _blnMeIsSource = true
    Public ReadOnly Property SourceColumn() As Integer
        Get
            Return SourceCell.ColumnIndex
        End Get
    End Property

    Private _blnRO As Boolean = False
    Public Property blnRO() As Boolean
        Get
            Return _blnRO
        End Get
        Set(ByVal value As Boolean)
            _blnRO = value
        End Set
    End Property

    'Event to enable dg as source of data.
    Protected Overridable Sub dgParent_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown

        'Dim p As Point = dgParent.PointToClient(New Point(e.X, e.Y))
        'In this case the coordinates are those of the grid.
        'Clicks = 1 is necessary to ensure that double click to filter still works.
        If blnRO = False And AllowDrop = True And e.Clicks = 1 Then

            'Store the source coordinates.
            SourceCell = HitTest(e.X, e.Y)
            If SourceCell.RowIndex >= 0 And SourceCell.ColumnIndex >= 0 Then

                'Do not allow PK to be source of data for dropping on to 
                'the same datagrid. However this could be allowed if PK was being moved to another control?
                If DragAndDropColumnIsASource(SourceCell.ColumnIndex) = True Then

                    'This will fail if new row.
                    Try

                        'Store the source control so that destination knows who sent.
                        _blnMeIsSource = True
                        DoDragDrop(Rows(SourceCell.RowIndex).Cells(SourceCell.ColumnIndex).Value, _
                               DragDropEffects.Copy Or DragDropEffects.Move)
                    Catch ex As Exception
                        _blnMeIsSource = False
                    End Try

                End If
            End If
        End If
    End Sub

    Private Sub dg_Leave(sender As System.Object, e As System.EventArgs) Handles Me.Leave
        _blnMeIsSource = False
    End Sub

    'Enable dg as destination of data.
    Protected Overridable Sub dgParent_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter

        If AllowDrop = True Then
            If (e.Data.GetDataPresent(DataFormats.Text)) Then
                If (e.KeyState And DragDropEffects.Copy) = DragDropEffects.Copy Then
                    e.Effect = DragDropEffects.Copy
                Else
                    e.Effect = DragDropEffects.Move
                End If
            End If
        End If
    End Sub

    Protected Overridable Sub dgParent_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop

        If AllowDrop = True Then

            Try

                'Here the coordinates are those of the form.
                Dim p As Point = PointToClient(New Point(e.X, e.Y))
                Dim info2 As DataGridView.HitTestInfo = HitTest(p.X, p.Y)

                'So allow if source was another object and dest is any field or if source is this grid then only allow if dest is not PK.
                If info2.RowIndex >= 0 And info2.ColumnIndex >= 0 And DragAndDropColumnIsADestination(info2.ColumnIndex) = True Then

                    Dim cCell As DataGridViewCell = Nothing
                    Dim blnIsAMove As Boolean = False
                    If _blnMeIsSource = False Then
                        DragAndDropUpdateRow(blnIsAMove, True, info2.RowIndex, info2.ColumnIndex, e.Data.GetData(DataFormats.Text))
                    Else

                        'Only allow if destination is other cell than source
                        If (info2.RowIndex <> SourceCell.RowIndex Or info2.ColumnIndex <> SourceCell.ColumnIndex) Then

                            'Remove the source cell value when moving.
                            If DragAndDropIsMove(SourceCell.ColumnIndex) Then
                                blnIsAMove = True
                                DragAndDropUpdateRow(blnIsAMove, False, SourceCell.RowIndex, SourceCell.ColumnIndex, e.Data.GetData(DataFormats.Text))
                                'Rows(SourceCell.RowIndex).Cells(SourceCell.ColumnIndex).Value = ""
                            End If
                            DragAndDropUpdateRow(blnIsAMove, True, info2.RowIndex, info2.ColumnIndex, e.Data.GetData(DataFormats.Text))
                        End If


                        'Is also possible to select the cell but is not necessary.
                        'cCell = dgParent.Rows(info2.RowIndex).Cells(info2.ColumnIndex)
                        'dgParent.CurrentCell = cCell
                        'dgParent.CurrentRow.Selected = True
                        'dgParent.CurrentCell.Value = e.Data.GetData(DataFormats.Text)
                    End If
                Else

                    'this is true if we drop onto an empty part of the grid.
                    DragAndDropInsertRow(e.Data.GetData(DataFormats.Text))

                End If

            Catch ex As Exception

            End Try
            'reset source field if the source was this DGV.
            _blnMeIsSource = False

        End If

    End Sub
    Protected Overridable Sub DragAndDropInsertRow(ByVal strValue As String)

    End Sub

    Protected Overridable Sub DragAndDropUpdateRow(ByVal blnIsAMove As Boolean, ByVal bnDestination As Boolean, ByVal RowIndex As Integer, ByVal ColumnIndex As Integer, ByVal strValue As String)
        Rows(RowIndex).Cells(ColumnIndex).Value = strValue
        DragAndDropUpdateAfter()
    End Sub

    ''' <summary>
    ''' Default behaviour is to update the complete table to save Source and Destination.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overridable Sub DragAndDropUpdateAfter()
        UpdateAllData()
    End Sub

    ''' <summary>
    ''' Default behaviour is that all columns in the grid may be a Source.
    ''' Override to get more control.
    ''' </summary>
    ''' <param name="iColumn"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overridable Function DragAndDropColumnIsASource(ByVal iColumn As Integer) As Boolean
        Return True
    End Function

    ''' <summary>
    ''' Default behaviour is that all columns may be a Destination.
    ''' Override to get more control over which columns may be a destination.
    ''' Use blnMeIsSource which is true if the Source column is from the same grid as the Destination.
    ''' If blnMeIsSource is true then use SourceColumn to identify the Source column.
    ''' </summary>
    ''' <param name="iColumnIndex"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overridable Function DragAndDropColumnIsADestination(ByVal iColumnIndex As Integer) As Boolean
        Return True
    End Function

    ''' <summary>
    ''' Default behaviour is to move from Source to Destination.
    ''' Set blnMove = false to copy from Source to Destination for all Source columns.
    ''' Get more control over which source columns are copied and which are moved by overriding this. 
    ''' </summary>
    ''' <param name="iColumnIndex"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overridable Function DragAndDropIsMove(ByVal iColumnIndex As Integer) As Boolean
        Return _blnMove
    End Function
#End Region
End Class
