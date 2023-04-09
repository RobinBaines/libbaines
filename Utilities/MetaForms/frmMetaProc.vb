'------------------------------------------------
'Name: Module frmMetaProc.vb
'Function: Show the procs defined in the database.
'Copyright Baines 2013. All rights reserved.
'Notes: 
'Modifications: 
'------------------------------------------------
'Imports Utilities
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
'Imports Microsoft.SqlServer.Management.SqlParser.Parser
''Imports Microsoft.Data.Schema.ScriptDom.Sql
Public Class frmMetaProc
    Dim vROUTINES As MetaData_v_INFORMATION_SCHEMA_ROUTINES
    Dim vReferenced As MetaData_v_referenced_objects
    Dim vReferencing As MetaData_v_referencing_objects

#Region "new"
    Public Sub New(ByVal tsb As ToolStripItem _
              , ByVal strSecurityName As String, ByVal _MainDefs As MainDefinitions)

        MyBase.New(tsb, strSecurityName, _MainDefs)
        InitializeComponent()


        vReferenced = New MetaData_v_referenced_objects(strSecurityName, Dm_sql_referenced_entitiesBindingSource, Me.dgReferenced, Dm_sql_referenced_entitiesTableAdapter, _
Me.MetaData, _
Me.components, _
MainDefs, True, Controls, Me, True)

        vReferencing = New MetaData_v_referencing_objects(strSecurityName, Dm_sql_referencing_entitiesBindingSource, Me.dgReferencing, Dm_sql_referencing_entitiesTableAdapter, _
Me.MetaData, _
Me.components, _
MainDefs, True, Controls, Me, True)



        vROUTINES = New MetaData_v_INFORMATION_SCHEMA_ROUTINES(strSecurityName, ROUTINESBindingSource, Me.dgROUTINES, ROUTINESTableAdapter, _
Me.MetaData, _
Me.components, _
MainDefs, True, Controls, Me, True)

        Me.SwitchOffPrintDetail()
        Me.SwitchOffPrint()
        Me.SwitchOffUpdate()
    End Sub

#End Region

#Region "Load"
    Protected Overrides Sub frmLoad(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles MyBase.Load
        MyBase.frmLoad(sender, e)
        Try
            tbROUTINES.WordWrap = False

            Me.tbROUTINES.Location = New Point(dgROUTINES.Location.X + dgROUTINES.Width + 10, dgROUTINES.Location.Y)
            bProcCopy.Location = New Point(dgROUTINES.Location.X + dgROUTINES.Width + 10, bProcCopy.Location.Y)
            vReferencing.AdjustPosition(vReferenced)
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
                vROUTINES.StoreRowIndexWithFocus()
                Me.ROUTINESTableAdapter.Fill(Me.MetaData.ROUTINES)
                vROUTINES.ResetFocusRow()
                blnAllowUpdate = True


                'tbROUTINES.SelectionColor = Color.Coral

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub dgROUTINES_RowEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgROUTINES.RowEnter
        Try
            If blnAllowUpdate = True Then

                Dm_sql_referenced_entitiesTableAdapter.Fill(Me.MetaData.dm_sql_referenced_entities, dgROUTINES.Rows(e.RowIndex).Cells("ROUTINE_SCHEMA").Value + "." + dgROUTINES.Rows(e.RowIndex).Cells("ROUTINE_NAME").Value)
                Dm_sql_referencing_entitiesTableAdapter.Fill(Me.MetaData.dm_sql_referencing_entities, dgROUTINES.Rows(e.RowIndex).Cells("ROUTINE_SCHEMA").Value + "." + dgROUTINES.Rows(e.RowIndex).Cells("ROUTINE_NAME").Value)

             
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "parse_sql"

    Private Sub tbROUTINES_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles tbROUTINES.TextChanged
        Try
            MainDefs.MainForm.SQLParser.ParseRTB(tbROUTINES)
        Catch ex As Exception

        End Try

    End Sub

#End Region

#Region "Scroll"
    Protected Overrides Sub frm_Layout(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles MyBase.Layout
        MyBase.frm_Layout(sender, e)
        If TestActiveMDIChild() = True Then
            vROUTINES.SetHeight(Me.ClientRectangle.Height - dgROUTINES.Location.Y - 140)
            tbROUTINES.Height = dgROUTINES.Height
            vReferenced.AdjustVerticalPosition(vROUTINES)
            vReferencing.AdjustVerticalPosition(vROUTINES)

            vReferenced.SetHeight(Me.ClientRectangle.Height)
            vReferencing.SetHeight(Me.ClientRectangle.Height)

            If dgReferencing.Location.X + dgReferencing.Width < tbROUTINES.Location.X Then
                Me.tbROUTINES.Height = Me.ClientRectangle.Height - tbROUTINES.Location.Y
            End If
        End If
    End Sub
#End Region

#Region "copy"
    Private Sub bProcCopy_Click(sender As Object, e As EventArgs) Handles bProcCopy.Click
        Me.Cursor = Cursors.WaitCursor
        Clipboard.SetText(Me.tbROUTINES.Text)
        Cursor = Cursors.Default
    End Sub
#End Region

    Private Sub dgReferenced_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgReferenced.RowHeaderMouseClick
        ResetFilter()
        If dgReferenced.Rows(e.RowIndex).Cells("referenced_schema_name").Value.Equals(System.DBNull.Value) Then
            vROUTINES.FindAndGotoRow("ROUTINE_NAME", dgReferenced.Rows(e.RowIndex).Cells("referenced_entity_name").Value)
        Else
            vROUTINES.FindAndGotoRow("COMPLETE_NAME", dgReferenced.Rows(e.RowIndex).Cells("referenced_schema_name").Value + "." + dgReferenced.Rows(e.RowIndex).Cells("referenced_entity_name").Value)
        End If


    End Sub

    Private Sub dgReferencing_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgReferencing.RowHeaderMouseClick
        ResetFilter()
        If dgReferencing.Rows(e.RowIndex).Cells("referencing_schema_name").Value.Equals(System.DBNull.Value) Then
            vROUTINES.FindAndGotoRow("ROUTINE_NAME", dgReferencing.Rows(e.RowIndex).Cells("referencing_entity_name").Value)
        Else
            vROUTINES.FindAndGotoRow("COMPLETE_NAME", dgReferencing.Rows(e.RowIndex).Cells("referencing_schema_name").Value + "." + dgReferencing.Rows(e.RowIndex).Cells("referencing_entity_name").Value)
        End If
    End Sub
End Class