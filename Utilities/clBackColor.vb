
'------------------------------------------------
'Name: Module clBackColor.vb
'Function: Cell formatting in datagridview controls.
'Copyright Baines 2008. All rights reserved.
'Notes:
'Modifications:
'------------------------------------------------
Imports System.Configuration
Imports System.Drawing
'Imports Utilities
Public Class clBackColor

    'use this if the cell being shown is a colour. Dont show the text.
    Public Shared Sub CellFormatting(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs)
        Try
            If e.ColumnIndex >= 0 Then
                If sender.Columns(e.ColumnIndex).Name.ToString().ToLower() = "backcolor" Then
                    If Not sender.Rows(e.RowIndex).Cells("backcolor").Value Is System.DBNull.Value Then
                        e.CellStyle.BackColor = Color.FromArgb(sender.Rows(e.RowIndex).Cells("backcolor").Value)
                        e.CellStyle.ForeColor = Color.FromArgb(sender.Rows(e.RowIndex).Cells("backcolor").Value)
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    'use a colour to set the background of the cell and set the text colour to black or white depending on the brightness of the background.
    'if the cell is selected before the background is set, happens when the form is opened, then it will not show until the selection is changed.
    'so the last clause sets the selection any way. 
    Public Shared Sub CellFormattingOther(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs, _
                                          strCellToColour As String, strCellToColourWith As String)
        Try
            If e.ColumnIndex >= 0 Then
                If sender.Columns(e.ColumnIndex).Name.ToString().ToLower() = strCellToColour Then
                    If Not sender.Rows(e.RowIndex).Cells(strCellToColourWith).Value Is System.DBNull.Value Then
                        Dim iValue As Integer = sender.Rows(e.RowIndex).Cells(strCellToColourWith).Value
                        e.CellStyle.BackColor = Color.FromArgb(iValue)
                        e.CellStyle.ForeColor = statics.GetTextColor(e.CellStyle.BackColor)

                        'set the selection colours too.
                        If e.RowIndex = sender.currentrow.index Then
                            sender.DefaultCellStyle.SelectionBackColor = e.CellStyle.BackColor
                            sender.DefaultCellStyle.SelectionForeColor = e.CellStyle.ForeColor
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Shared Sub CellFormattingUsingAnotherField(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs, _
                                      strCellToColour As String, strCellToColourWith As String)
        Try
            If e.ColumnIndex >= 0 Then
                If sender.Columns(e.ColumnIndex).Name.ToString().ToLower() = strCellToColour Then
                    If Not e.Value Is System.DBNull.Value Then
                        If e.Value.ToString.Length > 0 Then
                            If Not sender.Rows(e.RowIndex).Cells(strCellToColourWith).Value Is System.DBNull.Value Then
                                Dim iValue As Integer = sender.Rows(e.RowIndex).Cells(strCellToColourWith).Value
                                e.CellStyle.BackColor = Color.FromArgb(iValue)
                                e.CellStyle.ForeColor = statics.GetTextColor(e.CellStyle.BackColor)
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

End Class
