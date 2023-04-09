'------------------------------------------------
'Name: Module ToolStripCheckBox.vb
'Function: Make a checkbox for use in a toolstrip.
'Created Jan 2011.
'Notes: Use ToolStripControlHost as base class for a checkbox for a toolstrip.
'Funny that that is not available.
'Modifications:
'------------------------------------------------
Imports System.Windows.Forms
Public Class ToolStripCheckBox
    Inherits ToolStripControlHost
    Public Sub New()
        MyBase.New(New CheckBox)
    End Sub

    Private ReadOnly Property CheckBoxControl() As CheckBox
        Get
            Return CType(Control, CheckBox)
        End Get
    End Property

    Public Property Checked() As Boolean
        Get
            Return CheckBoxControl.Checked
        End Get
        Set(ByVal value As Boolean)
            CheckBoxControl.Checked = value
        End Set
    End Property

    Public Event CheckedChanged As EventHandler
    Private Sub CheckedChangedHandler(ByVal sender As Object, ByVal e As EventArgs)
        RaiseEvent CheckedChanged(Me, e)
    End Sub

    Protected Overrides Sub OnUnsubscribeControlEvents(ByVal control As System.Windows.Forms.Control)
        MyBase.OnUnsubscribeControlEvents(control)
        Dim checkBoxControl As CheckBox = CType(control, CheckBox)
        RemoveHandler checkBoxControl.CheckedChanged, AddressOf CheckedChangedHandler
    End Sub

    Protected Overrides Sub OnSubscribeControlEvents(ByVal control As System.Windows.Forms.Control)
        MyBase.OnSubscribeControlEvents(control)
        Dim checkBoxControl As CheckBox = CType(control, CheckBox)
        AddHandler checkBoxControl.CheckedChanged, AddressOf CheckedChangedHandler
    End Sub
End Class