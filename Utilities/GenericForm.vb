'------------------------------------------------
'Name: Module GenericForm.vb.
'Function: A child form for editing a single grid.
'GenericForm is derived from frmStandard and then adds a single dgvEnter datagridview.
'Copyright Robin Baines 2008. All rights reserved.
'Created April 2008.
'20120623 Changed Public Sub New(ByVal tsb As ToolStripButton to ToolStripItem to make more generic.
'------------------------------------------------
Imports System.Windows.Forms
Imports System.Drawing


'--------------------------------------------------------------
'--GenericForm Class
'----------------------------
Public Class GenericForm
    Inherits frmStandard
    Protected WithEvents dgParent As dgvEnter
    Protected WithEvents taParent As Object = Nothing
    Protected WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ContextMenuStrip2 As System.Windows.Forms.ContextMenuStrip
    Protected WithEvents vParent As dgColumns         'Generated object with columns and filter for a datagridview
    Public FilterFields As New Dictionary(Of String, String)
    Public WithEvents bsParent As System.Windows.Forms.BindingSource
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Dim blnLoaded As Boolean 'necessary because ShowDialog calls Load.
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.dgParent = New dgvEnter  'System.Windows.Forms.DataGridView
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.bsParent = New System.Windows.Forms.BindingSource(Me.components)
        CType(Me.dgParent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.bsParent, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgParent
        '
        Me.dgParent.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgParent.ContextMenuStrip = Me.ContextMenuStrip1
        Me.dgParent.Location = New System.Drawing.Point(3, 61)
        Me.dgParent.Margin = New System.Windows.Forms.Padding(4)
        Me.dgParent.Name = "dgParent"
        Me.dgParent.RowTemplate.Height = 24
        Me.dgParent.Size = New System.Drawing.Size(657, 709)
        Me.dgParent.TabIndex = 1
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(61, 4)
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(61, 4)
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'GenericForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(1093, 1077)
        Me.Controls.Add(Me.dgParent)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "GenericForm"
        Me.Text = "Generic"
        Me.Controls.SetChildIndex(Me.dgParent, 0)
        CType(Me.dgParent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.bsParent, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Protected Overridable Sub EndInitializeComponent()

        ' CType(Me.bsParent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgParent, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

#Region "New"
    Public ReadOnly Property _components() As System.ComponentModel.IContainer
        Get
            Return components
        End Get
    End Property

    Protected ut As New Utilities
    Dim toolStripProgressBar As ToolStripProgressBar
    Dim progressStatusStrip As StatusStrip
    Dim toolStripStatusLabel As ToolStripStatusLabel
    Dim flp As FlowLayoutPanel
    Dim blnFullSize As Boolean

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub New_init(ByVal tsb As ToolStripItem _
    , ByVal strSecurityName As String _
    , ByVal _MainDefs As MainDefinitions _
    , ByVal Fields As Dictionary(Of String, String), ByVal _blnFullSize As Boolean, _
    ByVal bRO As Boolean)

        If bRO = True Then blnRO = True
        InitializeComponent()
        blnLoaded = False
        ResizeTheControls()
        EndInitializeComponent()
        blnFullSize = _blnFullSize
        Me.SwitchOffPrintDetail()

        'Store the filter values.
        CopyFilterField(Fields)
        If blnRO = True Then SwitchOffUpdate()

    End Sub

    Public Sub New(ByVal tsb As ToolStripItem _
    , ByVal strSecurityName As String _
    , ByVal _MainDefs As MainDefinitions _
    , ByVal Fields As Dictionary(Of String, String), ByVal _blnFullSize As Boolean, _
    ByVal bRO As Boolean)

        MyBase.New(tsb, strSecurityName, _MainDefs)
        New_init(tsb, strSecurityName, _MainDefs, Fields, _blnFullSize, bRO)
    End Sub

    '20130123 Added this constructor to differentiate between form name and security name. Is needed when opening a master data
    'form as child tsm of a parent tsm. Parent tsm is the security name and the form text should be Form.
    Public Sub New(ByVal tsb As ToolStripItem _
        , _strForm As String, ByVal strSecurityName As String _
        , ByVal _MainDefs As MainDefinitions _
        , ByVal Fields As Dictionary(Of String, String), ByVal _blnFullSize As Boolean, _
        ByVal bRO As Boolean)

        MyBase.New(tsb, _strForm, strSecurityName, _MainDefs)
        New_init(tsb, strSecurityName, _MainDefs, Fields, _blnFullSize, bRO)

    End Sub

    Public Sub CopyFilterField(ByVal Fields As Dictionary(Of String, String))
        FilterFields.Clear()
        If Not Fields Is Nothing Then
            For Each kvp As KeyValuePair(Of String, String) In Fields
                FilterFields.Add(kvp.Key, kvp.Value)
            Next
        End If
    End Sub

    Public Sub SetFilter(ByVal Fields As Dictionary(Of String, String))
        CopyFilterField(Fields)
        vParent.ColumnDoubleClick(FilterFields)
    End Sub

    Protected Overridable Sub CreateFilterBoxes()

    End Sub

    Protected Overridable Sub ResizeTheControls()

    End Sub

#End Region

#Region "Load"
    Private Sub GenericLoad(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If blnLoaded = False Then
            blnAllowUpdate = False
            If vGridsCount = 0 Then
                Adjustcolumns()
                AdjustFilterBoxes()
            End If
            ConfigureContextStrip()
            If Not vParent Is Nothing Then
                vParent.dgResize(Me.Size.Width, Me.GetScrollState(ScrollStateVScrollVisible))
            End If
            Me.SetBindingNavigatorSource(bsParent)
            blnLoaded = True
        End If
        If Not vParent Is Nothing Then
            vParent.ResetFilter()
            vParent.SelectFirstTb()
        End If
        If blnFullSize = False Then AdjustFormHeight()
        blnAllowUpdate = True
        AdjustColourOfButtons()
        FillTableAdapter()
        HelpTextPosition()
    End Sub

    'Override this to do nothing.
    Public Overrides Sub FormIsActivated()
       
    End Sub
    Protected Overridable Sub Adjustcolumns()
        If Not vParent Is Nothing Then
            vParent.Adjustcolumns(True)
        End If

    End Sub

    Protected Overridable Sub AdjustFilterBoxes()
        If Not vParent Is Nothing Then
            vParent.AdjustFilterBoxes()
        End If
    End Sub

    Protected Sub AdjustFormHeight()

        'Set the height if only 1 record is visible.
        If Me.dgParent.RowCount <= 20 Then
            BindingNavigator.Visible = False
            Dim iDelta = dgParent.Height - (Me.dgParent.RowCount * 25 + 50)   '100
            dgParent.Height = dgParent.Height - iDelta
            Me.Height = dgParent.Height + 80 'Me.Height - iDelta
        End If
    End Sub

    Protected Overridable Sub ConfigureContextStrip()
        'the contextstrip1 is bound to dgParent and is shown automatically on right mouse click.
    End Sub

    Private Sub dgParent_RowEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgParent.RowEnter
        vParent.RowEnter(Me.blnAllowUpdate, sender, e)
    End Sub

    Private Sub dgParent_CellFormatting(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dgParent.CellFormatting
        vParent.CellFormatting(sender, e)
    End Sub

#End Region

#Region "Filter"
    Public Overrides Sub ResetFilter()
        MyBase.ResetFilter()
        vParent.ResetFilter()
    End Sub

    Public Overrides Sub FilterFromOtherForm(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        MyBase.FilterFromOtherForm(sender, e)
        vParent.ColumnDoubleClick(sender, e)
    End Sub

    Public Function strGetFieldName(ByVal strName As String) As String()
        'Get the field name from a string showing Table, {Field}
        'Return the second name after the '.'
        Dim strRet(10) As String

        'Remove the table name.
        Dim iFields As Integer = 0
        Dim strT As String
        Dim i As Integer = strName.IndexOf(".")
        If i > 0 Then
            strT = strName.Substring(i + 1)
            Do While iFields < (strRet.Length - 1)
                i = strT.IndexOf(".")
                If i > 0 Then
                    strRet(iFields) = strT.Substring(0, i)
                    strT = strT.Substring(i + 1)
                Else
                    strRet(iFields) = strT
                    iFields = iFields + 1
                    Exit Do
                End If
                iFields = iFields + 1
            Loop
        End If
        strRet(iFields) = ""
        Return strRet
    End Function

    Public Sub ShowForm(ByVal pParent As Form, ByVal strForm As String, ByVal MainDefs As MainDefinitions, ByVal Fields As Dictionary(Of String, String))
        Dim strHeader = MainDefs.strGetTableText(strForm)
        ShowForm(pParent, strHeader, strForm, MainDefs, Fields)
    End Sub

    Public Overridable Sub ShowForm(ByVal pParent As Form, ByVal strHeader As String, ByVal strForm As String, ByVal MainDefs As MainDefinitions, ByVal Fields As Dictionary(Of String, String))
    End Sub

    Protected Sub OpenFilteredForm(ByVal sender As System.Object)
        'The sender is a ToolStripMenuItem with Text equal to the Child Form name and
        'Name = ChildFormName.{Fields} where the Fields are those which should be filtered in the child form.
        Dim Fields As New Dictionary(Of String, String)

        'Put the filter fields into an array of strings.
        Dim strPKColumns As String() = strGetFieldName(sender.Tag)
        Dim strFieldName As String() = strGetFieldName(sender.Name)
        Dim iField As Integer = 0

        'Build a dictionary with Field Name and Field value.
        Do While strFieldName(iField).Length And strPKColumns(iField).Length
            Fields.Add(strFieldName(iField), dgParent.CurrentRow.Cells(strPKColumns(iField)).Value.ToString())
            iField = iField + 1
        Loop

        'Open the form showing the hierarchy in the header name.
        Dim strForm As String = sender.Name
        strForm = strForm.Substring(0, strForm.IndexOf("."))
        ShowForm(Me.MdiParent, Me.Text + "." + sender.Text, strForm, MainDefs, Fields)
    End Sub

#End Region

#Region "Scroll"

    'HelpTextBox always under the Parent grid.
    Public Overrides Sub HelpTextPosition()
        Dim iheight = Me.ClientRectangle.Height - HelpTextBox.Height - 20
        HelpTextBox.Location = New Point(dgParent.Location.X, iheight)
    End Sub

    ' Private Sub frmProject_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
    Protected Overrides Sub frm_Layout(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles MyBase.Layout
        If Not vParent Is Nothing Then
            If TestActiveMDIChild() = True Then
                vParent.dgResize(sender.Width, Me.GetScrollState(ScrollStateVScrollVisible))
                Dim iheight As Integer = 0
                If HelpTextBox.Visible = True Then
                    iheight = HelpTextBox.Height + 20
                End If
                If Not vParent Is Nothing Then
                    vParent.SetHeight(Me.ClientRectangle.Height - iheight)
                    HelpTextPosition()
                End If

            End If
        End If
    End Sub

    Dim iScrollPosition As Integer = 0
    Private Sub dgParent_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles dgParent.Scroll
        If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then

            'Adjust filter visibility and position.

            iScrollPosition = e.NewValue
            vParent.AdjustFilterBoxes(e.NewValue)
        End If
    End Sub

    Private Sub dgParent_ColumnWidthChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgParent.ColumnWidthChanged

        If blnAllowUpdate = True Then

            'Adjust filter visibility and position.
            vParent.AdjustFilterBoxes(iScrollPosition)
        End If
    End Sub

    Private Sub frm_SizeChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.SizeChanged
        If Not vParent Is Nothing Then
            vParent.Fillout(DataGridViewAutoSizeColumnMode.NotSet)
        End If
    End Sub

#End Region

#Region "Print"

    Protected Overrides Sub tsbPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Print overview.
        Dim strHeader = Me.Name
        Dim strFilename = strHeader.Replace(" ", "_").Replace(".", "_")
        PrintToExcel(strFilename, strFilename, dgParent)
    End Sub

#End Region

    Protected Overridable Sub dgParent_CellBeginEdit(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgParent.CellBeginEdit

    End Sub

    Protected Overridable Sub dgParent_CellValidating(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgParent.CellValidating

    End Sub

    Protected Overridable Sub dgParent_CellValidated(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgParent.CellValidated

    End Sub

    Private Sub TToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub

    Protected Overridable Sub RefreshCombos()
    End Sub

    Private Sub tsbRefreshCombos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        RefreshCombos()
        dgParent.CancelEdit()
    End Sub

End Class
