'------------------------------------------------
'Name: Module frmMetaView.vb
'Function: Show the views defined in the database.
'Copyright Baines 2013. All rights reserved.
'Notes: 
'Modifications: 
'------------------------------------------------

Imports System.Windows.Forms
Imports System.Drawing

Public Class frmMetaView
    Dim vParent As ADHOCDataSet_v_all_views
    Dim vChild1 As MetaData_v_all_views_column
    Dim vReferenced As MetaData_v_referenced_objects
    Dim vReferencing As MetaData_v_referencing_objects

#Region "new"
    Public Sub New(ByVal tsb As ToolStripItem _
              , ByVal strSecurityName As String, ByVal _MainDefs As MainDefinitions)

        MyBase.New(tsb, strSecurityName, _MainDefs)
        InitializeComponent()
        vParent = New ADHOCDataSet_v_all_views(strSecurityName, v_all_viewsBindingSource, dgParent, v_all_viewsTableAdapter, _
Me.MetaData, _
Me.components, _
MainDefs, True, Controls, Me, True)

        vReferenced = New MetaData_v_referenced_objects(strSecurityName, Dm_sql_referenced_entitiesBindingSource, Me.dgReferenced, Dm_sql_referenced_entitiesTableAdapter, _
Me.MetaData, _
Me.components, _
MainDefs, True, Controls, Me, True)

        vReferencing = New MetaData_v_referencing_objects(strSecurityName, Dm_sql_referencing_entitiesBindingSource, Me.dgReferencing, Dm_sql_referencing_entitiesTableAdapter, _
Me.MetaData, _
Me.components, _
MainDefs, True, Controls, Me, True)


        vChild1 = New MetaData_v_all_views_column(strSecurityName, V_all_views_columnBindingSource, dgChild1, V_all_views_columnTableAdapter, _
Me.MetaData, _
Me.components, _
MainDefs, True, Controls, Me, True)



        Me.SwitchOffPrintDetail()
        Me.SwitchOffPrint()
        Me.SwitchOffUpdate()
        Me.HelpTextBox.Visible = False
    End Sub

#End Region

#Region "Load"
    Protected Overrides Sub frmLoad(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles MyBase.Load
        MyBase.frmLoad(sender, e)
        Try
            Me.tbView.WordWrap = False
            vChild1.AdjustPosition(vParent)
            tbView.Location = New Point(dgChild1.Location.X + dgChild1.Width + 10, dgChild1.Location.Y)
            bViewCopy.Location = New Point(dgChild1.Location.X + dgChild1.Width + 10, bViewCopy.Location.Y)

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
                vParent.StoreRowIndexWithFocus()
                Me.v_all_viewsTableAdapter.Fill(Me.MetaData.v_all_views)
                vParent.ResetFocusRow()
                blnAllowUpdate = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub dgParent_RowEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgParent.RowEnter
        Try
            If blnAllowUpdate = True Then
                Me.V_all_views_columnTableAdapter.FillBy(Me.MetaData.v_all_views_column, dgParent.Rows(e.RowIndex).Cells("TABLE_NAME").Value, dgParent.Rows(e.RowIndex).Cells("TABLE_SCHEMA").Value)
                Try
                    Dm_sql_referenced_entitiesTableAdapter.Fill(Me.MetaData.dm_sql_referenced_entities, dgParent.Rows(e.RowIndex).Cells("TABLE_SCHEMA").Value + "." + dgParent.Rows(e.RowIndex).Cells("TABLE_NAME").Value)
                Catch ex As Exception
                    MsgBox("CHECK A referenced entity does not exist and View is invalid. " + ex.Message)
                End Try
                Try
                    Dm_sql_referencing_entitiesTableAdapter.Fill(Me.MetaData.dm_sql_referencing_entities, dgParent.Rows(e.RowIndex).Cells("TABLE_SCHEMA").Value + "." + dgParent.Rows(e.RowIndex).Cells("TABLE_NAME").Value)
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region
#Region "parse_sql"
    Private Sub tbView_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles tbView.TextChanged
        'TestStatics.ParseRTB(tbView)
        Try
            MainDefs.MainForm.SQLParser.ParseRTB(tbView)
        Catch ex As Exception
        End Try

    End Sub
#End Region
#Region "Scroll"
    Protected Overrides Sub frm_Layout(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles MyBase.Layout
        MyBase.frm_Layout(sender, e)
        If TestActiveMDIChild() = True Then
            vParent.SetHeight(Me.ClientRectangle.Height - dgParent.Location.Y - 140)
            vChild1.SetHeight(Me.ClientRectangle.Height - dgChild1.Location.Y - 140)
            tbView.Height = dgParent.Height
            vReferenced.AdjustVerticalPosition(vParent)
            vReferencing.AdjustVerticalPosition(vParent)
            ' vROUTINES.SetHeight(Me.ClientRectangle.Height)
            vReferenced.SetHeight(Me.ClientRectangle.Height)
            vReferencing.SetHeight(Me.ClientRectangle.Height)

            If dgReferencing.Location.X + dgReferencing.Width < tbView.Location.X Then
                tbView.Height = Me.ClientRectangle.Height - tbView.Location.Y
            End If


        End If
    End Sub
#End Region

#Region "copy"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles bViewCopy.Click
        Me.Cursor = Cursors.WaitCursor
        Clipboard.SetText(tbView.Text)
        Cursor = Cursors.Default
    End Sub
#End Region

    Private Sub dgReferenced_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgReferenced.RowHeaderMouseClick
        Try
            Dim strFind As String
            If dgReferenced.Rows(e.RowIndex).Cells("referenced_schema_name").Value.Equals(System.DBNull.Value) Then
                strFind = dgReferenced.Rows(e.RowIndex).Cells("referenced_entity_name").Value
                ResetFilter()
                vParent.FindAndGotoRow("TABLE_NAME", strFind)
            Else
                strFind = dgReferenced.Rows(e.RowIndex).Cells("referenced_schema_name").Value + "." + dgReferenced.Rows(e.RowIndex).Cells("referenced_entity_name").Value
                ResetFilter()
                vParent.FindAndGotoRow("COMPLETE_NAME", strFind)
            End If
        Catch ex As Exception
        End Try

    End Sub

    ' ResetFilter() changes the Rows() array so store the find string first.
    Private Sub dgReferencing_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgReferencing.RowHeaderMouseClick
        Try
            Dim strFind As String
            If dgReferencing.Rows(e.RowIndex).Cells("referencing_schema_name").Value.Equals(System.DBNull.Value) Then
                strFind = dgReferencing.Rows(e.RowIndex).Cells("referencing_entity_name").Value
                ResetFilter()
                vParent.FindAndGotoRow("TABLE_NAME", strFind)
            Else
                strFind = dgReferencing.Rows(e.RowIndex).Cells("referencing_schema_name").Value + "." + dgReferencing.Rows(e.RowIndex).Cells("referencing_entity_name").Value
                ResetFilter()
                vParent.FindAndGotoRow("COMPLETE_NAME", strFind)
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class

