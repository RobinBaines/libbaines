'------------------------------------------------
'Name: Class dgvPlus.vb.
'Function: Overrides the default Enter key behaviour changing it to move right.
'Copyright Robin Baines 2010. All rights reserved.
'------------------------------------------------
Imports System.Windows.Forms
Imports System.Windows.Forms.DataGridViewAdvancedBorderStyle
Imports System.Drawing
Public Class dgvPlus
    Inherits DataGridView
    Friend WithEvents __ta As Object = Nothing
    Public Property ta() As Object
        Get
            Return __ta
        End Get
        Set(ByVal value As Object)
            __ta = value
        End Set
    End Property
    Public Sub New()
        MyBase.New()
    End Sub
    Protected Overridable Sub dg_CellValidating(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles Me.CellValidating
    End Sub
    Private Sub dg_CellValidated(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Me.CellValidated
        _blnDirty = True
    End Sub
    Protected Overrides Function ProcessDialogKey(ByVal keyData As Keys) As Boolean
        If keyData = Keys.Enter Then
            Dim CurrentIndex As Integer = Me.CurrentCell.ColumnIndex
            Dim blnRet As Boolean = Me.ProcessRightKey(keyData)

            'When editing it is important that the cell does move to another one so if this is the last cell
            'send it to the left.
            If CurrentIndex = Me.CurrentCell.ColumnIndex Then
                Me.ProcessLeftKey(keyData)
                If CurrentIndex = Me.CurrentCell.ColumnIndex Then

                    'This occurs if there is only one field.
                    If Me.EndEdit() = True Then
                        _blnDirty = True
                    End If
                End If
            End If
            Return blnRet
        End If
        Return MyBase.ProcessDialogKey(keyData)
    End Function
    Protected Overrides Function ProcessDataGridViewKey(ByVal e As KeyEventArgs) As Boolean
        If e.KeyCode = Keys.Enter Then
            Dim blnRet = Me.ProcessRightKey(e.KeyData)
            Return blnRet
        End If

        '
        If e.KeyCode = Keys.Escape Then
            _blnDirty = False
        End If

        Return MyBase.ProcessDataGridViewKey(e)
    End Function

    Private Sub dg_EditingControlShowing(ByVal sender As Object, _
    ByVal e As DataGridViewEditingControlShowingEventArgs) _
    Handles Me.EditingControlShowing
        ' e.CellStyle.BackColor = Color.Aquamarine
    End Sub
    Private _blnDirty As Boolean
    Public Property blnDirty() As Boolean
        Get
            Return _blnDirty
        End Get
        Set(ByVal value As Boolean)
            _blnDirty = value
        End Set
    End Property
    Public Sub UpdateData()
        If Me.EndEdit() = True Then

            'If Endedit did something.
            Dim bs As BindingSource = CType(Me.DataSource, BindingSource)
            bs.EndEdit()
            Dim sD As System.Data.DataRowView
            sD = bs.Current

            'Used the table adapter stored in the dgvEnter object to update the Row.
            If Not Me.ta Is Nothing Then
                Me.ta.update(sD.Row)
            End If
            _blnDirty = False
            Me.Refresh()
        End If
    End Sub
    Protected Overrides Sub OnRowPostPaint(ByVal e As DataGridViewRowPostPaintEventArgs)

        '//this method overrides the DataGridView's RowPostPaint event 
        If Not Me.CurrentRow Is Nothing Then
            If e.RowIndex = Me.CurrentRow.Index And _blnDirty = True Then

                Dim strRowNumber As String = (e.RowIndex + 1).ToString()
                Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)

                '//draw the row number string on the current row header cell using
                '//the brush defined above and the DataGridView's default font
                If Me.RowHeadersWidth > 20 Then

                    Dim rect As New RectangleF(e.RowBounds.Location.X + Me.RowHeadersWidth - 20, e.RowBounds.Location.Y + 5, 10, e.RowBounds.Height - 10)
                    e.Graphics.FillRectangle(Brushes.AliceBlue, rect)
                End If
                ' e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

            End If
        End If
        MyBase.OnRowPostPaint(e)
    End Sub

End Class
