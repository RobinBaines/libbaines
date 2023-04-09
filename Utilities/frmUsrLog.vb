'------------------------------------------------
'Name: Module frmUsrLog.vb
'Function: 
'Copyright Robin Baines 2010. All rights reserved.
'Created March 2010.
'Purpose: 
'Notes: 
'Modifications: 
'------------------------------------------------
Imports System.Windows.Forms
Imports System.Drawing
Public Class frmUsrLog
    Dim vParent As TheDataSet_v_usr_log
    Dim vChild1 As TheDataSet_m_usr_log
    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub
    Public Sub New(ByVal tsb As ToolStripItem _
               , ByVal strSecurityName As String, ByVal _MainDefs As MainDefinitions)

        MyBase.New(tsb, strSecurityName, _MainDefs)
        InitializeComponent()
        vParent = New TheDataSet_v_usr_log(strSecurityName, V_usr_logBindingSource, dgParent, V_usr_logTableAdapter, _
              Me.TheDataSet, _
              Me.components, _
              MainDefs, True, Controls, Me, True)
        vChild1 = New TheDataSet_m_usr_log(strSecurityName, M_usr_logBindingSource, dgChild1, M_usr_logTableAdapter, _
              Me.TheDataSet, _
              Me.components, _
              MainDefs, True, Controls, Me, True)
        SetBindingNavigatorSource(M_usr_logBindingSource)
        Me.SwitchOffPrintDetail()
        Me.SwitchOffUpdate()
    End Sub
    Protected Overrides Sub frmLoad(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.frmLoad(sender, e)
        vChild1.AdjustPosition(vParent)
        FillTableAdapter()
    End Sub
    Protected Overrides Sub FillTableAdapter()
        MyBase.FillTableAdapter()
        vParent.StoreRowIndexWithFocus()
        Me.V_usr_logTableAdapter.Fill(Me.TheDataSet.v_usr_log)
        vParent.ResetFocusRow()
    End Sub
    Private Sub dgParent_RowEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgParent.RowEnter
        Try
            If Not dgParent Is Nothing Then
                Dim strT As String = dgParent.Rows(e.RowIndex).Cells("usr").Value
                strT = dgParent.Rows(e.RowIndex).Cells("app").Value
                Me.M_usr_logTableAdapter.FillByUsr(Me.TheDataSet.m_usr_log, dgParent.Rows(e.RowIndex).Cells("usr").Value, dgParent.Rows(e.RowIndex).Cells("app").Value)
            End If
        Catch ex As Exception
        End Try
    End Sub
#Region "Filter"
    Public Overrides Sub FilterFromOtherForm(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        MyBase.FilterFromOtherForm(sender, e)
        vParent.ColumnDoubleClick(sender, e)
    End Sub
#End Region
#Region "Print"
    Protected Overrides Sub tsbPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim strHeader = Me.Name
        Dim strFilename = strHeader.Replace(" ", "_").Replace(".", "_")
        PrintToExcel(strFilename, strFilename, dgParent)
    End Sub
#End Region
#Region "Scroll"
    Protected Overrides Sub frm_Layout(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LayoutEventArgs)
        MyBase.frm_Layout(sender, e)
        If TestActiveMDIChild() = True Then
            If Not vParent Is Nothing Then
                vParent.SetHeight(Me.ClientRectangle.Height)
            End If
            If Not vChild1 Is Nothing Then
                vChild1.SetHeight(Me.ClientRectangle.Height)
            End If
        End If
    End Sub
#End Region
End Class
