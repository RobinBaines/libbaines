'------------------------------------------------
'Name: Module frmHelpText.vb
'Function: HelpText form.
'Copyright Robin Baines 2022. All rights reserved.
'Notes: 
'Modifications: 
'------------------------------------------------
Imports Utilities
Imports System
Imports System.Windows.Forms
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Threading
Imports System.Drawing
Imports ExcelInterface.XMLExcelInterface
Public Class frmHelpText
    Dim vParent1 As m_helptext
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
        vParent1 = New m_helptext(strSecurityName, M_form_helptextBindingSource, M_form_helptextDataGridView, M_form_helptextTableAdapter, _
            Me.TheDataSet, _
            Me.components, _
            MainDefs, blnRO, Controls, Me, True)
        SetBindingNavigatorSource(M_form_helptextBindingSource)
        'vParent1.CreateFilterBoxes(Me.Controls)

        Me.SwitchOffPrintDetail()
        Me.SwitchOffPrint()
        'SetPrintDetail(statics.get_txt_header("Print Colors"))
        'SwitchOffNavigator()
        'Me.SwitchOffRefresh()
        Me.SwitchOffUpdate()


    End Sub

#End Region
#Region "Load"
    Protected Overrides Sub frmLoad(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.frmLoad(sender, e)

        M_form_helptextDataGridView.AllowUserToAddRows = False
        M_form_helptextDataGridView.AllowUserToDeleteRows = False
        blnAllowUpdate = True
        FillTableAdapter()
    End Sub
    Protected Overrides Sub FillTableAdapter()
        MyBase.FillTableAdapter()
        If blnAllowUpdate = True Then
            vParent1.StoreRowIndexWithFocus()
            Me.M_form_helptextTableAdapter.Fill(Me.TheDataSet.m_form_helptext)
            vParent1.ResetFocusRow()
            HelpTextPosition()
        End If
    End Sub
#End Region

#Region "Scroll"

    Protected Overrides Sub frm_Layout(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LayoutEventArgs)
        MyBase.frm_Layout(sender, e)
        Try
            If Not vParent1 Is Nothing Then
                Dim iheight As Integer = 0
                If HelpTextBox.Visible Then
                    iheight = HelpTextBox.Height + 20
                End If
                vParent1.SetHeight(Me.ClientRectangle.Height - iheight)
                HelpTextPosition()
            End If

        Catch ex As Exception

        End Try
    End Sub
#End Region

    '#Region "Scroll_Resize"
    '    'And dont forget the filter boxes on the colours.
    '    'Private Sub frm_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
    '    Protected Overrides Sub frm_Layout(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles MyBase.Layout
    '        If TestActiveMDIChild() = True Then

    '            If Not M_form_helptextDataGridView Is Nothing Then
    '                Dim iDelta = 10 + statics.GroupBoxRelativeVerticalLocation

    '                M_form_helptextDataGridView.Height = (Me.Height - 120 - iDelta) / 2
    '                Dim l As Point
    '                l = Me.M_form_helptextDataGridView.Location
    '                l.Y = 50
    '                Me.M_form_helptextDataGridView.Location = l
    '                l = vParent1.gbForFiltersGroupBox.Location
    '                l.Y = Me.M_form_helptextDataGridView.Location.Y - statics.GroupBoxRelativeVerticalLocation
    '                Me.vParent1.gbForFiltersGroupBox.Location = l
    '            End If
    '        End If
    '    End Sub
    '#End Region

End Class