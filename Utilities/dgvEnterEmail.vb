'------------------------------------------------
'Name: Module dgvEnterEmail.vb
'Function: a datagridview object derived from dgventer with validation of email string.
'Copyright Robin Baines 2012. All rights reserved.
'Notes: 
'Modifications: 
'------------------------------------------------
Imports System.Windows.Forms
Imports System.Windows.Forms.DataGridViewAdvancedBorderStyle
Imports System.Drawing
'Imports Utilities

Public Class dgvEnterEmail
    Inherits dgvEnter
    Dim dgRowAltered As Boolean = False

    Public Sub New()
        MyBase.New()
    End Sub

    Private Sub ValidateEmail(e As System.Windows.Forms.DataGridViewCellValidatingEventArgs)
        Me.Rows(e.RowIndex).ErrorText = ""
        Dim util As New RegexUtilities()
        If e.FormattedValue.length > 0 Then
            If util.IsValidEmail(e.FormattedValue) = False Then
                e.Cancel = True
                Me.Rows(e.RowIndex).ErrorText = statics.get_txt_header("Not a valid email address.", "User advice in dgvEnterEmail.", "User information")
            End If
        End If
    End Sub

    Protected Overrides Sub dg_CellValidating(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles Me.CellValidating
        If Not IsDBNull(e.FormattedValue) Then
            If e.ColumnIndex = Me.Columns("email").Index Then
                ValidateEmail(e)
            End If
            If Not Me.Columns("cc_email") Is Nothing Then
                If e.ColumnIndex = Me.Columns("cc_email").Index Then
                    ValidateEmail(e)
                End If
            End If

        End If
    End Sub

End Class


