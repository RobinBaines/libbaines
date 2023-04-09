'------------------------------------------------
'Name: Module frmAppLog.vb
'Function: Show the application log.
'Copyright Robin Baines 2010. All rights reserved.
'Notes: 
'Modifications: 
'------------------------------------------------
Imports Utilities
Imports System
Imports System.Windows.Forms
Imports System.Data.SqlClient
Public Class frmAppLog
    Dim vParent As m_app_log

#Region "New"

    Public Sub New()

        MyBase.New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

    End Sub
    Public Sub New(ByVal tsb As ToolStripItem _
               , ByVal strSecurityName As String, ByVal _MainDefs As MainDefinitions)

        MyBase.New(tsb, strSecurityName, _MainDefs)
        InitializeComponent()
        vParent = New m_app_log(strSecurityName, M_app_logBindingSource, dgParent, M_app_logTableAdapter, _
            Me.TheDataSet, _
            Me.components, _
            MainDefs, True, Controls, Me, True)

        SetBindingNavigatorSource(M_app_logBindingSource)
        Me.SwitchOffPrintDetail()
        Me.SwitchOffUpdate()
    End Sub

#End Region

#Region "Load"

    Protected Overrides Sub frmLoad(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles MyBase.Load
        MyBase.frmLoad(sender, e)
        blnAllowUpdate = True
        FillTableAdapter()
    End Sub
    Protected Overrides Sub FillTableAdapter()
        MyBase.FillTableAdapter()
        vParent.StoreRowIndexWithFocus()
        Me.M_app_logTableAdapter.Fill(Me.TheDataSet.m_app_log)
        vParent.ResetFocusRow()

    End Sub
#End Region

#Region "Scroll_Resize"
    'The datagrids are re-sized when the form re-sizes. But this causes problems if the re-size fires when the window is not the ActiveMDIChild.
    'This occurs if the Ctrl-tab combination is used to cycle through the windows of the application followed by an Alt.
    'This was also a problem when the form was re-writing when semaphore fired.
    'Solution is only to re-size when the form is the ActiveMDIChild.
    'Tried also to check on the windowstate so Resize occurs if the windowstate is not maximised.
    'But it appears that the windowstate is Normal if is maximized but is not the ActiveMDIChild.
    'Private Sub frm_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
    Protected Overrides Sub frm_Layout(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LayoutEventArgs) 'Handles MyBase.Layout
        MyBase.frm_Layout(sender, e)
        If TestActiveMDIChild() = True Then
            If Not vParent Is Nothing Then
                vParent.SetHeight(Me.ClientRectangle.Height) ' dgParent.Height = Me.Height - 40 - dgParent.Location.Y
            End If
        End If
    End Sub
#End Region
    
End Class