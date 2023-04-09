'------------------------------------------------
'Name: Module frmStandard.vb
'Function: The form class which is inherited when creating new forms.
'Copyright Robin Baines 2008. All rights reserved.
'Created Jan 2008.
'Notes: 
'Modifications: Also translate control text in flowlayoutpanel.
'20150815 added parameter lsemaphore when distributing event because semaphore has altered.
'20200102 added pr = Nothing to close XMLExcelInterface.
'20200915 added Protected Overridable Sub SetDialogStartPosition(). Default is 40, 40.
'20221014 added help text.
'------------------------------------------------
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Xml
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Drawing
Imports System.Runtime
Imports ExcelInterface.XMLExcelInterface

Imports Utilities
Public Class frmStandard

    Public vGrids As New List(Of dgColumns)

    'Set to True to call RefreshTheForm() from FormIsActivated()
    Public blnRefreshIsNeeded As Boolean = False

    Public MainDefs As MainDefinitions
    Protected tsbCalling As ToolStripItem
    Private _blnRO As Boolean
    Protected iInitialFormHeight As Integer = 1000

    Dim strSecurityName As String
    Public ReadOnly Property SecurityName() As String
        Get
            Return strSecurityName
        End Get
    End Property

    Dim strForm As String
    Public ReadOnly Property FormName() As String
        Get
            Return strForm
        End Get
    End Property

    Dim _Enable_ChkLevel_On_ChildControls As Boolean = False


    ''' <summary>
    '''  '20120516 Set to allow fine tuning of controls in the form.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Enable_ChkLevel_On_ChildControls() As Boolean
        Get
            Return _Enable_ChkLevel_On_ChildControls
        End Get
        Set(ByVal value As Boolean)
            _Enable_ChkLevel_On_ChildControls = value
        End Set
    End Property

    Public Property blnRO() As Boolean
        Get
            Return _blnRO
        End Get
        Set(ByVal value As Boolean)
            _blnRO = value
        End Set
    End Property

    Public ReadOnly Property vGridsCount() As Integer
        Get
            Return vGrids.Count
        End Get
    End Property

    Private _blnAllowUpdate As Boolean = False
    Public Property blnAllowUpdate() As Boolean
        Get
            Return _blnAllowUpdate
        End Get
        Set(ByVal value As Boolean)
            _blnAllowUpdate = value
        End Set
    End Property

#Region "EnterTimer"
    Public Overridable Sub RowEnterTimerEvent(dg As dgvEnter)

    End Sub
#End Region

#Region "AdjustForm"
    Protected Sub BindingNavigatorInvisible()
        Me.BindingNavigator.Visible = False
    End Sub
    Protected Sub BindingNavigatorVisibility(ByVal blnVisible As Boolean)
        Me.BindingNavigator.Visible = blnVisible
    End Sub
    Protected Sub SetBindingNavigatorSource(ByVal bs As BindingSource)
        Me.BindingNavigator.BindingSource = bs
    End Sub
    Protected Sub BindingNavigatorAdd(ByVal tsb As ToolStripItem)   ' ToolStripButton)
        Me.BindingNavigator.Items.Add(tsb)
        tsb.ToolTipText = tsb.Text
        tsb.Text = statics.get_txt_header(tsb.ToolTipText)
    End Sub
    Protected Sub SwitchOffNavigator()

        BindingNavigatorPositionItem.Visible = False
        BindingNavigatorMoveFirstItem.Visible = False
        BindingNavigatorMovePreviousItem.Visible = False
        BindingNavigatorCountItem.Visible = False
        BindingNavigatorMoveNextItem.Visible = False
        BindingNavigatorMoveLastItem.Visible = False
        BindingNavigatorSeparator.Visible = False
        BindingNavigatorSeparator1.Visible = False
        BindingNavigatorSeparator2.Visible = False

    End Sub

    Protected Sub SwitchOffRefresh()
        Me.tsbRefresh.Visible = False
    End Sub

    Protected Sub SwitchOffUpdate()
        Me.tsbUpdate.Visible = False
    End Sub

    Protected Sub SwitchOffPrint()
        Me.tsbPrint.Visible = False
    End Sub

    Protected Sub SwitchOffPrintDetail()
        Me.tsbPrintdetail.Visible = False
    End Sub

    Protected Sub SwitchOffFormName()
        tbFormName.Visible = False
    End Sub

    Protected Sub SetPrintDetail(ByVal strText As String)
        Me.tsbPrintdetail.Text = strText
    End Sub
    Protected Sub SetRefresh(ByVal strText As String)
        Me.tsbRefresh.Text = strText
    End Sub

    Protected Function SetRefreshColor(ByVal cC As Color) As Color
        Dim CurrentColor = Me.tsbRefresh.BackColor
        Me.tsbRefresh.BackColor = cC
        Return CurrentColor
    End Function

    Protected Function GetRefreshColor() As Color
        Return Me.tsbRefresh.BackColor
    End Function

    Protected Function GetRefresh() As String
        Return Me.tsbRefresh.Text
    End Function

    'This only does the standard buttons. The text of ones being added is adjusted as they added above.
    Private Sub AdjustTextOfButtons()
        For Each c As ToolStripItem In Me.BindingNavigator.Items
            Dim c_tsb As System.Windows.Forms.ToolStripButton
            c_tsb = TryCast(c, ToolStripButton)
            If Not c_tsb Is Nothing Then
                c_tsb.ToolTipText = c_tsb.Text
                c_tsb.Text = statics.get_txt_header(c_tsb.ToolTipText)
            End If
        Next
    End Sub

    Public Sub AdjustColourOfButtons()
        Dim cBackColor As Color = Color.AliceBlue
        For Each c As ToolStripItem In Me.BindingNavigator.Items
            If c.Visible = True And c.Name.Contains("BindingNavigator") = False Then
                c.BackColor = cBackColor
                If cBackColor = Color.Lavender Then
                    cBackColor = Color.AliceBlue
                Else
                    cBackColor = Color.Lavender
                End If
            End If
        Next
    End Sub

#End Region

#Region "HelpText"
    Dim blnHelpTextVertical As Boolean = False
    Public Property HelpTextVertical() As Boolean
        Get
            Return blnHelpTextVertical
        End Get
        Set(value As Boolean)
            blnHelpTextVertical = value
        End Set
    End Property

    Public ReadOnly Property HelpTextHeight() As Integer
        Get
            If HelpTextBox.Visible Then
                Return HelpTextBox.Height + 20
            End If
            Return 0
        End Get
    End Property

    Public Overridable Sub HelpTextPosition()
        '20221014
        If HelpTextBox.Visible = True Then
            Dim LeftPos = 100
            Dim iWidth = 0
            Dim iHeight = 0
            For Each vGrid As dgColumns In vGrids
                If vGrid.dg.Location.X + vGrid.dg.Width > iWidth Then
                    iWidth = vGrid.dg.Location.X + vGrid.dg.Width
                End If

                If vGrid.dg.Location.Y + vGrid.dg.Height > iHeight Then
                    iHeight = vGrid.dg.Location.Y + vGrid.dg.Height
                End If

                If vGrid.dg.Location.X < LeftPos Or LeftPos = 0 Then
                    LeftPos = vGrid.dg.Location.X
                End If

            Next

            For Each ctl As Control In Controls
                Dim tb As TextBox
                tb = TryCast(ctl, TextBox)
                If Not tb Is Nothing Then
                    If tb.Name <> "HelpTextBox" Then
                        If tb.Location.X + tb.Width > iWidth And tb.Name <> "HelpTextBox" Then
                            iWidth = tb.Location.X + tb.Width
                        End If
                        If ctl.Location.Y + ctl.Height > iHeight Then
                            iHeight = ctl.Location.Y + ctl.Height
                        End If
                    End If
                Else
                    Dim tl As Label
                    tl = TryCast(ctl, Label)
                    If Not tl Is Nothing Then
                        If tl.Location.X + tl.Width > iWidth Then
                            iWidth = tl.Location.X + tl.Width
                        End If
                        If ctl.Location.Y + ctl.Height > iHeight Then
                            iHeight = ctl.Location.Y + ctl.Height
                        End If
                    Else
                        Dim cb As ComboBox
                        cb = TryCast(ctl, ComboBox)
                        If Not cb Is Nothing Then
                            If cb.Location.X + cb.Width > iWidth Then
                                iWidth = cb.Location.X + cb.Width
                            End If
                            If ctl.Location.Y + ctl.Height > iHeight Then
                                iHeight = ctl.Location.Y + ctl.Height
                            End If
                        Else
                            Dim ck As CheckBox
                            ck = TryCast(ctl, CheckBox)
                            If Not ck Is Nothing Then
                                If ck.Location.X + ck.Width > iWidth Then
                                    iWidth = ck.Location.X + ck.Width
                                End If
                                If ctl.Location.Y + ctl.Height > iHeight Then
                                    iHeight = ctl.Location.Y + ctl.Height
                                End If
                            End If

                        End If
                    End If
                End If
            Next

            If (iWidth > iHeight And iWidth > 1499) Or Me.Modal = True Then
                HelpTextBox.Location = New Point(LeftPos, iHeight + 20)
                blnHelpTextVertical = True
            Else
                HelpTextBox.Location = New Point(iWidth + 20, HelpTextBox.Location.Y)
                blnHelpTextVertical = False
            End If

        End If
    End Sub

    Protected Overridable Sub GetHelpText()
        Me.HelpTextBox.Visible = False
        Try
            Dim m_form_helptextTableAdapter = New TheDataSetTableAdapters.m_form_helptextTableAdapter
            m_form_helptextTableAdapter.Connection.ConnectionString = MainDefs.GetConnectionString()
            HelpTextBox.Text = m_form_helptextTableAdapter.GetHelpText(Me.SecurityName)
            If HelpTextBox.Text.Length > 0 Then
                Me.HelpTextBox.Visible = True
            End If
        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "New"
    Public Sub New()

        MyBase.New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

    End Sub

    '20141219 Moved here from the designer code.
    'Form overrides dispose to clean up the component list.
    '<System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If

            For Each vGrid As dgColumns In vGrids
                vGrid.ADispose()
            Next
            vGrids.Clear()

        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '20200915 added Protected Overridable Sub SetDialogStartPosition(). Default is 40, 40.
    Protected Overridable Sub SetDialogStartPosition()
        StartPosition = FormStartPosition.Manual
        Location = New Point(40, 40)
    End Sub

    Private Sub New_init(ByVal tsb As ToolStripItem _
               , _strForm As String, ByVal _strSecurityName As String, ByVal _MainDefs As MainDefinitions, blnChildDialog As Boolean)


        '20120717 remove fast key prefix from security names.
        strForm = _strForm.Replace("&", "")
        strSecurityName = _strSecurityName.Replace("&", "")
        If Not tsb Is Nothing Then
            BindingNavigatorMoveFirstItem.BackColor = tsb.BackColor
            Me.BindingNavigatorMoveLastItem.BackColor = tsb.BackColor
            Me.BindingNavigatorMovePreviousItem.BackColor = tsb.BackColor
            Me.BindingNavigatorMoveNextItem.BackColor = tsb.BackColor
        End If

        MainDefs = _MainDefs
        tsbCalling = tsb
        AdjustTextOfButtons()

        '20100217 Set in blnCheckLevel
        'blnRO = True
        If blnChildDialog Then
            statics.blnCheckLevel(strSecurityName, blnRO, True, True)
        Else
            statics.blnCheckLevel(strSecurityName, blnRO)
        End If


        Me.AutoScroll = True
        SelectedFont(tsb)
        iInitialFormHeight = Me.Height + 30

        Me.KeyPreview = True

        '20200915 added Protected Overridable Sub SetDialogStartPosition(). Default is 40, 40.
        If Me.Modal = True Then
            SetDialogStartPosition()
        End If

        GetHelpText()

    End Sub

    Public Sub New(ByVal tsb As ToolStripItem _
           , ByVal _strSecurityName As String, ByVal _MainDefs As MainDefinitions)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        New_init(tsb, _strSecurityName, _strSecurityName, _MainDefs, False)
    End Sub

    'added this constructor to differentiate between form name and security name. Is needed when opening a master data
    'form as child tsm of a parent tsm. Parent tsm is the security name and the form text should be Form.
    Public Sub New(ByVal tsb As ToolStripItem _
               , _strForm As String, ByVal _strSecurityName As String, ByVal _MainDefs As MainDefinitions)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        New_init(tsb, _strForm, _strSecurityName, _MainDefs, False)
    End Sub

    'added this constructor for Dialogs which are children of another form.
    'Children means the Level check is usually linked to the read only status of the parent and that the Form name does not appear in the
    'list of forms when the Group is defined (menu_entry = 1). This analogous to Master menu entries which follow their Parent menu
    'which is why use of menu_entry is justified.
    Public Sub New(ByVal _strForm As String, ByVal _MainDefs As MainDefinitions)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        New_init(Nothing, _strForm, _strForm, _MainDefs, True)
    End Sub

    '20120702 Set and reset the parent.
    Protected Sub SelectedFont(ByVal tsb As ToolStripItem)
        If Not tsb Is Nothing Then
            Dim SlFont As Font = New System.Drawing.Font("Tahoma", 8, FontStyle.Bold Or FontStyle.Regular)

            'get the parent. When the parent font is set it propagates to all the children whihc is not what we want.
            Dim ts As ToolStrip = tsb.GetCurrentParent()
            If Not ts Is Nothing Then
                Dim tsdmi As ToolStripDropDownMenu = TryCast(ts, ToolStripDropDownMenu)
                If Not tsdmi Is Nothing Then
                    'store the existing fonts of the children.
                    Dim SlFonts As New List(Of Font)
                    For Each tsi As ToolStripItem In ts.Items
                        Dim tsmi As ToolStripMenuItem = TryCast(tsi, ToolStripMenuItem)
                        If Not tsmi Is Nothing Then
                            If tsmi.Equals(tsb) Then
                                SlFonts.Add(SlFont)
                            Else
                                SlFonts.Add(tsmi.Font)
                            End If
                        End If
                    Next

                    'set the font of the parent. This propagates to the children.
                    If Not tsdmi.OwnerItem Is Nothing Then
                        tsdmi.OwnerItem.Font = SlFont


                        'Reset the children back to the correct font.
                        Dim iInc As Integer = 0
                        For Each tsi As ToolStripItem In ts.Items
                            Dim tsmi As ToolStripMenuItem = TryCast(tsi, ToolStripMenuItem)
                            If Not tsmi Is Nothing Then
                                tsmi.Font = SlFonts(iInc)
                                iInc += 1
                            End If
                        Next
                    End If
                Else
                    tsb.Font = SlFont
                End If
            End If
        End If
    End Sub

    Public Sub DeSelectedFont(ByVal tsb As ToolStripItem)
        Try

            'Make font of Button normal instead of bold.
            Dim SelectFont As System.Drawing.Font
            If Not tsb Is Nothing Then
                SelectFont = New System.Drawing.Font("Tahoma", 8, FontStyle.Regular)
                tsb.Font = SelectFont
                'Reset the parent.

                Dim ts As ToolStrip = tsb.GetCurrentParent()
                If Not ts Is Nothing Then
                    Dim tsdmi As ToolStripDropDownMenu = TryCast(ts, ToolStripDropDownMenu)
                    Dim blnAllClosed = True
                    If Not tsdmi Is Nothing Then
                        If Not tsdmi.OwnerItem Is Nothing Then
                            For Each tsi As ToolStripItem In ts.Items
                                Dim tsmi As ToolStripMenuItem = TryCast(tsi, ToolStripMenuItem)
                                If Not tsmi Is Nothing Then
                                    If tsmi.Font.Bold = True Then
                                        blnAllClosed = False
                                        Exit For
                                    End If
                                End If
                            Next
                            If blnAllClosed = True Then
                                tsdmi.OwnerItem.Font = New System.Drawing.Font("Tahoma", 8, FontStyle.Regular)
                            End If
                        End If
                    Else
                        tsb.Font = SelectFont
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "Load"
    'use to fit a datagridview which is the only one vertically.
    'call from form MyBase.Resize.
    Public Sub ResizeGrid(ByVal dg As DataGridView)
        Dim iAdjustable = dg.Height
        If iInitialFormHeight <> 0 Then
            Dim dFactor As Double = Me.Height / iInitialFormHeight
            dg.Height = dg.Height * dFactor
        End If

    End Sub

    '20120319 This added because RAP re-uses forms for dialogs and Load gets call on Dialog.Show.
    Dim blnLoaded As Boolean = False
    Protected Overridable Sub frmLoad(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        AdjustColourOfButtons()
        If blnRO = True Then
            ReadOnlyTextBoxes(True)
        End If

        '20110606 This is needed if Layout event is used to resize datagrids.
        Me.AutoScroll = False

        '20110418. This is needed to force maximized when a form is loaded.
        'Also forces a re-size.
        If Me.Modal = False Then
            Me.WindowState = FormWindowState.Maximized
        Else
            Me.WindowState = FormWindowState.Normal
        End If

        '20120210
        For Each vGrid As dgColumns In vGrids
            If blnLoaded = False Then
                vGrid.CreateFilterBoxes()
                vGrid.Adjustcolumns(True)
            End If
            vGrid.AdjustFilterBoxes()
        Next

        '20120418 alter tooltips and label text of controls.
        If Not strForm Is Nothing Then
            If strForm.Length > 0 Then
                Me.Text = statics.get_txt_header(strForm, "form name", strForm)
            End If
        End If

        '20120911 Translate all text. Do this before CreateTooltips whihc uses binding to find the labels.
        '20120516 This functionality allows fine tuning of enabling of controls in the form.
        If _Enable_ChkLevel_On_ChildControls = True Then
            Me.CreateTooltips(Me.components)
        Else
            Try
                TranslateAllText(Me.components)
                If blnRO = True Then
                    DisableButtons()
                    ReadOnlyTextBoxes(True)
                    ReadOnlyComboBoxes(True)
                    ReadOnlyCheckBoxes(True)
                    ReadOnlyListBoxes(True)
                    ReadOnlyNumericUpDown(True)
                    ReadOnlyGroupBox(True)
                End If
            Catch ex As Exception

            End Try
        End If

        '20161126 Show auto scroll for laptops.
        If Screen.AllScreens.Length = 1 And Screen.PrimaryScreen.Bounds.Width < 2000 Then
            AutoScroll = True
        End If

        HelpTextPosition()

        blnLoaded = True
    End Sub

    '20120210
    Protected Overridable Sub frm_Layout(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles MyBase.Layout
        Me.SuspendLayout()
        For Each vGrid As dgColumns In vGrids
            vGrid.dgResize(sender.Width, Me.GetScrollState(ScrollStateVScrollVisible))
        Next
        Me.ResumeLayout()
    End Sub

    '20190115 Added frm_Scroll handler to resize grids when user releases the horizontal scroll bar.
    Private Sub frm_Scroll(sender As Object, e As ScrollEventArgs) Handles MyBase.Scroll
        If e.ScrollOrientation = ScrollOrientation.HorizontalScroll Then
            If e.Type <> ScrollEventType.ThumbTrack Then
                Me.SuspendLayout()
                For Each vGrid As dgColumns In vGrids
                    vGrid.dgResize(sender.Width, Me.GetScrollState(ScrollStateVScrollVisible))
                Next
                Me.ResumeLayout()
            End If
        End If
    End Sub

    'The datagrids are re-sized when the form re-sizes. But this causes problems if the re-size fires when the window is not the ActiveMDIChild.
    'This occurs if the Ctrl-tab combination is used to cycle through the windows of the application followed by an Alt.
    'Not clear why that should be but think Alt returns focus and forces a re-size.
    'Updating the form (resize or redraw) was also a problem when the form was re-writing when semaphore fired.
    'Solution is only to resize/redraw when the form is the ActiveMDIChild.
    'Tried also to check on the windowstate so Resize occurs if the windowstate is not maximised.
    'But it appears that the windowstate is Normal if is maximized but is not the ActiveMDIChild.
    Protected Function TestActiveMDIChild() As Boolean
        'Return True
        If Not MainDefs Is Nothing Then
            If MainDefs.MainForm.ActiveMdiChild Is Nothing Then Return False
            If MainDefs.MainForm.ActiveMdiChild.Name = Me.Name Then
                Return True
            End If
        End If
        Return False
    End Function

    ''' <summary>
    ''' Is called from MainFrom RefreshAll()
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RefreshTheForm()
        FillTableAdapter()
    End Sub

    Public blnHasFocus As Boolean
    '20130131 Underline the tsb name of the form with the focus.
    Private Sub frm_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If Not tsbCalling Is Nothing Then
            Dim SlFont As Font = New System.Drawing.Font(tsbCalling.Font, FontStyle.Bold Or FontStyle.Underline)
            tsbCalling.Font = SlFont
            tbFormName.Text = tsbCalling.Text
            tbFormName.Font = SlFont
            blnHasFocus = True
        End If
    End Sub

    Private Sub frmCycles_Deactivate(sender As Object, e As EventArgs) Handles MyBase.Deactivate
        If Not tsbCalling Is Nothing Then
            Dim SlFont As Font

            'Check whether bold or not as form may be being closed. 
            If tsbCalling.Font.Bold = True Then
                SlFont = New System.Drawing.Font(tsbCalling.Font, FontStyle.Bold Or FontStyle.Regular)
            Else
                SlFont = New System.Drawing.Font(tsbCalling.Font, FontStyle.Regular)
            End If
            tsbCalling.Font = SlFont
            blnHasFocus = False
        End If
    End Sub

    'Is called if a form is activated.
    Public Overridable Sub FormIsActivated()
        If blnRefreshIsNeeded = True Then
            RefreshTheForm()
            blnRefreshIsNeeded = False
        End If
    End Sub

    Public Overridable Sub FormIsNotActivated()

    End Sub

    'Is called if Database has requested a refresh.
    'Set blnRefreshIsNeeded to true if NoUpdate is checked so that data can be reloaded 
    'when NoUpdate is unchecked.
    ''' 20150815 added parameter lsemaphore
    Public Overridable Sub RefreshTheForm(ByVal activeForm As Form, ByVal strApp As String, ByVal strTble As String, ByVal lsemaphore As Long)

    End Sub

    Protected Sub RefreshAll()
        Dim frm As MainForm
        frm = CType(ParentForm, MainForm)
        If Not frm Is Nothing Then
            frm.RefreshAll()
        End If
    End Sub

    Public Overridable Sub FillTableAdapterPublic()
        FillTableAdapter()
    End Sub

    Protected Overridable Sub FillTableAdapter()

    End Sub

    'Called when update button is clicked. Override to endedit and write data.
    Protected Overridable Sub UpdateData()
        Dim dgv As dgvEnter

        'Look for dgvEnter controls.
        For Each c As Control In Controls
            dgv = TryCast(c, dgvEnter)
            If Not dgv Is Nothing Then
                Try
                    dgv.UpdateData()
                Catch ex As Exception
                End Try
            End If
        Next
    End Sub

    Protected Overridable Sub fStatus_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        'De select the font in the main form.
        DeSelectedFont(tsbCalling)
    End Sub


    ''' <summary>
    ''' 20091130 RPB added ResetTheCell.
    ''' Use these to reset the cell after the parent has finished loading.
    ''' </summary>
    ''' <param name="dg"></param>
    ''' <param name="cTheCell"></param>
    ''' <remarks></remarks>
    Protected Sub ResetTheCell(ByVal dg As DataGridView, ByVal cTheCell As DataGridViewCell)
        Try
            If Not cTheCell Is Nothing Then
                If cTheCell.Visible = True Then
                    dg.CurrentCell = cTheCell
                    dg.CurrentRow.Selected = True
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Protected Sub ReadOnlyTextBoxes(ByVal bRO As Boolean)
        For Each ctl As Control In Controls
            Dim tb As TextBox
            tb = TryCast(ctl, TextBox)
            If Not tb Is Nothing Then
                tb.ReadOnly = bRO
            End If
        Next
    End Sub

    Protected Sub ReadOnlyComboBoxes(ByVal bRO As Boolean)
        For Each ctl As Control In Controls
            Dim tb As ComboBox
            tb = TryCast(ctl, ComboBox)
            If Not tb Is Nothing Then
                tb.Enabled = Not bRO
            End If
        Next
    End Sub

    Protected Sub ReadOnlyCheckBoxes(ByVal bRO As Boolean)
        For Each ctl As Control In Controls
            Dim tb As CheckBox
            tb = TryCast(ctl, CheckBox)
            If Not tb Is Nothing Then
                tb.Enabled = Not bRO
            End If
        Next
    End Sub

    Protected Sub ReadOnlyListBoxes(ByVal bRO As Boolean)
        For Each ctl As Control In Controls
            Dim tb As ListBox
            tb = TryCast(ctl, ListBox)
            If Not tb Is Nothing Then
                tb.Enabled = Not bRO
            End If
        Next
    End Sub

    Protected Sub ReadOnlyNumericUpDown(ByVal bRO As Boolean)
        For Each ctl As Control In Controls
            Dim tb As NumericUpDown
            tb = TryCast(ctl, NumericUpDown)
            If Not tb Is Nothing Then
                tb.ReadOnly = bRO
            End If
        Next
    End Sub

    Protected Sub ReadOnlyGroupBox(ByVal bRO As Boolean)
        For Each ctl As Control In Controls
            Dim tb As GroupBox
            tb = TryCast(ctl, GroupBox)
            If Not tb Is Nothing Then
                If Not tb.Name = statics.FilterGroupBoxName Then
                    tb.Enabled = Not bRO
                End If

            End If
        Next
    End Sub
#End Region

#Region "GridManagement"
    Public Sub RegisterGrid(ByVal vGrid As dgColumns)
        vGrids.Add(vGrid)
    End Sub


#End Region

#Region "CreateTsb"
    Public Function CreateTsb(ByVal strName As String, ByVal strText As String, ByVal blnAddToToolstrip As Boolean) As ToolStripButton
        Return CreateTsb(strName, strText, blnAddToToolstrip, False)
    End Function
    Public Function CreateTsb(ByVal strName As String, ByVal strText As String, ByVal blnAddToToolstrip As Boolean, ByVal iWidth As Integer) As ToolStripButton
        Return CreateTsb(strName, strText, blnAddToToolstrip, False, iWidth)
    End Function
    Public Function CreateTsb(ByVal strName As String, ByVal strText As String, ByVal blnAddToToolstrip As Boolean, ByVal blnAutoSize As Boolean) As ToolStripButton
        Return CreateTsb(strName, strText, blnAddToToolstrip, blnAutoSize, 120)
    End Function
    'create a new tsb and optionally add to the binding navigator
    Public Function CreateTsb(ByVal strName As String, ByVal strText As String, ByVal blnAddToToolstrip As Boolean, ByVal blnAutoSize As Boolean _
    , ByVal iWidth As Integer) As ToolStripButton
        Dim tsb As ToolStripButton = New System.Windows.Forms.ToolStripButton
        tsb.Size = New System.Drawing.Size(iWidth, 25)
        tsb.AutoSize = blnAutoSize
        tsb.BackColor = System.Drawing.Color.Silver
        tsb.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        tsb.ImageTransparentColor = System.Drawing.Color.Magenta
        tsb.Name = strName

        tsb.Text = statics.get_txt_header(strText, "Tool strip button", strSecurityName)
        tsb.ToolTipText = strText
        If blnAddToToolstrip Then
            Me.BindingNavigator.Items.Add(tsb)
        End If
        Return tsb
    End Function
    Public Function CreateTsbCheckBox(ByVal strName As String, ByVal strText As String, ByVal blnAddToToolstrip As Boolean, ByVal blnAutoSize As Boolean _
   , ByVal iWidth As Integer) As ToolStripCheckBox
        Dim tsb As ToolStripCheckBox = New ToolStripCheckBox
        tsb.Size = New System.Drawing.Size(iWidth, 25)
        tsb.AutoSize = blnAutoSize
        tsb.BackColor = System.Drawing.Color.Silver
        tsb.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        tsb.ImageTransparentColor = System.Drawing.Color.Magenta
        tsb.Name = strName

        tsb.Text = statics.get_txt_header(strText, "Tool strip button", strSecurityName)
        tsb.ToolTipText = strText
        If blnAddToToolstrip Then
            Me.BindingNavigator.Items.Add(tsb)
        End If
        Return tsb
    End Function
#End Region

#Region "Focus"

    Protected Overridable Sub frm_MouseClick()
        Me.ScrollControlIntoView(Me.BindingNavigator)
    End Sub
    '20100910 If user clicks the form scroll to the top. 
    'Scenario is user selects a grid which causes the form to scroll up to show that grid.
    'User can now click the form surface to go to the top instead of having to scroll upwards.
    Private Sub frmStandard_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseClick
        frm_MouseClick()
    End Sub
#End Region

#Region "Filter"
    Public Sub BroadcastFilter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim frm As MainForm
        frm = CType(ParentForm, MainForm)
        frm.BroadcastFilter(sender, e)
    End Sub
    Public Overridable Sub ResetFilter()
        For Each vGrid As dgColumns In vGrids
            vGrid.ResetFilter()
        Next
    End Sub
    Public Overridable Sub FilterFromOtherForm(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
    End Sub
#End Region

#Region "Print"

    'wrapper to use My.Settings.XMLTemplate
    Protected Function strGetXMLtemplate() As String
        Return My.Settings.XMLTemplate
    End Function

    Public Sub PrintExcel(strFileName As String, dg As dgvEnter)
        PrintToExcel(strFileName, dg)
    End Sub

    Protected Sub PrintToExcel(ByVal strHeader As String, _
       ByVal dg As DataGridView)
        Dim strFilename = strHeader.Replace(" ", "_").Replace(".", "_").Replace("\", "_").Replace("/", "_")
        PrintToExcel(strFilename, strFilename, dg)
    End Sub
    Protected Sub PrintToExcel(ByVal strFilename As String, _
        ByVal strHeader As String, _
                ByVal dg As DataGridView)

        PrintToExcel(strFilename, strFilename, "", dg)
    End Sub

    '20161102 Added parameter strPath
    Protected Sub PrintToExcel(ByVal strPath As String, _
                               ByVal strFilename As String, _
      ByVal strHeader As String, _
      ByVal strFooter As String, _
      ByVal dg As DataGridView)

        Dim cur = Me.Cursor
        Try
            Me.Cursor = Cursors.WaitCursor

            Dim pr As New ExcelInterface.XMLExcelInterface("")  '"" is the Network path so we go local.
            pr.OpenExcelBook(strPath, strFilename, False, My.Settings.XMLTemplate, False)

            '20091005 Sheetname should not contain . / \ etc and not longer than 30.
            Dim strSheetName As String = strFilename
            If strSheetName.Length > 30 Then strSheetName = strSheetName.Substring(0, 30)
            pr.NewSheet(strSheetName, _
                "&amp;LUsing data from " + strFilename + "." + "&amp;CPrinted on &amp;D &amp;T. " + "&amp;RPage &amp;P of &amp;N", _
                True, strFooter)
            pr.WriteDataGrid(dg, MainDefs.DONOTPRINT, False, 0, False)
            pr.CloseExcelBook()
            NAR(pr) ' = Nothing
        Catch ex As Exception
            MsgBox("Could not create report." & ex.Message)
        Finally
            Me.Cursor = cur
        End Try
    End Sub

    Protected Sub PrintToExcel(ByVal strFilename As String, _
        ByVal strHeader As String, _
        ByVal strFooter As String, _
        ByVal dg As DataGridView)

        'PrintToExcel(ExcelInterface.Paths.Local, strFilename, strHeader, strFooter, dg)
        Dim cur = Me.Cursor
        Try
            Me.Cursor = Cursors.WaitCursor

            Dim pr As New ExcelInterface.XMLExcelInterface("")  '"" is the Network path so we go local.
            pr.OpenExcelBook(ExcelInterface.Paths.Local, "", strFilename, False, My.Settings.XMLTemplate)

            '20091005 Sheetname should not contain . / \ etc and not longer than 30.
            Dim strSheetName As String = strFilename
            If strSheetName.Length > 30 Then strSheetName = strSheetName.Substring(0, 30)
            pr.NewSheet(strSheetName, _
                "&amp;LUsing data from " + strFilename + "." + "&amp;CPrinted on &amp;D &amp;T. " + "&amp;RPage &amp;P of &amp;N", _
                True, strFooter)
            pr.WriteDataGrid(dg, MainDefs.DONOTPRINT, False, 0, False)
            pr.CloseExcelBook()
            NAR(pr) ' = Nothing

            'Runtime.GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce
            GC.Collect(2, GCCollectionMode.Forced)

        Catch ex As Exception
            MsgBox("Could not create report." & ex.Message)
        Finally
            Me.Cursor = cur
        End Try
    End Sub
    ' Print parent and 2 children
    Protected Sub PrintDetailToExcel(ByVal strHeader As String, ByVal dgParent As DataGridView, _
            ByVal dgChild1 As DataGridView, ByVal dgChild2 As DataGridView)
        Dim cur = Me.Cursor
        Try
            Me.Cursor = Cursors.WaitCursor

            Dim pr As New ExcelInterface.XMLExcelInterface("")  '"" is the Network path so we go local.
            PrintDetailToExcel(pr, strHeader.Replace(" ", "_").Replace(".", "_"), strHeader, "", dgParent, dgChild1)
            pr.WriteStringToExcel("#", ExcelInterface.ExcelStringFormats.Bold10)
            pr.WriteDataGrid(dgChild2, MainDefs.DONOTPRINT, False, 0, False, False, True)
            pr.CloseExcelBook()
            NAR(pr) ' = Nothing
        Catch ex As Exception
            MsgBox("Could not create report." & ex.Message)
        Finally
            Me.Cursor = cur
        End Try
    End Sub



    ''' <summary>
    ''' Print parent and 1 child
    ''' </summary>
    ''' <param name="strHeader"></param>
    ''' <param name="dgParent"></param>
    ''' <param name="dgChild1"></param>
    ''' <remarks></remarks>
    Protected Sub PrintDetailToExcel(ByVal strHeader As String, ByVal dgParent As DataGridView _
    , ByVal dgChild1 As DataGridView)
        Dim cur = Me.Cursor
        Try
            Me.Cursor = Cursors.WaitCursor
            Dim pr As New ExcelInterface.XMLExcelInterface("")  '"" is the Network path so we go local.
            PrintDetailToExcel(pr, strHeader.Replace(" ", "_").Replace(".", "_").Replace("\", "_").Replace("/", "_"), _
                    strHeader, "", dgParent, dgChild1)
            pr.CloseExcelBook()
            NAR(pr) ' = Nothing
        Catch ex As Exception
            MsgBox("Could not create report." & ex.Message)
        Finally
            Me.Cursor = cur
        End Try

    End Sub
    Protected Sub PrintDetailToExcel(ByVal strHeader As String, ByVal strFooter As String, ByVal dgParent As DataGridView _
    , ByVal dgChild1 As DataGridView)
        Dim cur = Me.Cursor
        Try
            Me.Cursor = Cursors.WaitCursor
            Dim pr As New ExcelInterface.XMLExcelInterface("")  '"" is the Network path so we go local.
            PrintDetailToExcel(pr, strHeader.Replace(" ", "_").Replace(".", "_").Replace("\", "_").Replace("/", "_"), _
                    strHeader, strFooter, dgParent, dgChild1)
            pr.CloseExcelBook()
            NAR(pr) ' = Nothing
        Catch ex As Exception
            MsgBox("Could not create report." & ex.Message)
        Finally
            Me.Cursor = cur
        End Try

    End Sub

    Public Sub NAR(ByRef o As Object)

        'MOD RPB 21st May 2007. Change to ByRef to ensure the object variable is reset.
        'See http://support.microsoft.com/default.aspx?scid=KB;EN-US;q317109
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch
        Finally
            o = Nothing
        End Try
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub PrintDetailToExcel(ByVal pr As ExcelInterface.XMLExcelInterface, _
            ByVal strFilename As String, _
            ByVal strHeader As String, ByVal strFooter As String, ByVal dgParent As DataGridView, _
            ByVal dgChild1 As DataGridView)

        'Template is a string and not a file reference.
        pr.OpenExcelBook(ExcelInterface.Paths.Local, "", strFilename, False, My.Settings.XMLTemplate)

        '20091005 Sheetname should not contain . / \ etc and not longer than 30.
        Dim strSheetName As String = strFilename
        If strSheetName.Length > 30 Then strSheetName = strSheetName.Substring(0, 30)
        pr.NewSheet(strSheetName, _
            "&amp;LUsing data from " + strFilename + "." + "&amp;CPrinted on &amp;D &amp;T. " + "&amp;RPage &amp;P of &amp;N", _
            True, strFooter)
        pr.WriteColumnWidths(dgParent, MainDefs.DONOTPRINT, False, 0)
        pr.WriteStringToExcel(strHeader, ExcelInterface.ExcelStringFormats.Bold10)
        pr.WriteStringToExcel("#", ExcelInterface.ExcelStringFormats.Bold10)
        pr.WriteDataGrid(dgParent, MainDefs.DONOTPRINT, True, 0, False, False, True)
        pr.WriteStringToExcel("#", ExcelInterface.ExcelStringFormats.Bold10)
        pr.WriteDataGrid(dgChild1, MainDefs.DONOTPRINT, False, 0, False, False, True)
        'pr.CloseExcelBook()
    End Sub
    Protected Overridable Sub tsbPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbPrint.Click
    End Sub
    Protected Overridable Sub tsbPrintdetail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbPrintdetail.Click
    End Sub
#End Region

#Region "Buttons"
    'Also  F10 see keybord handling
    Protected Overridable Sub tsbRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbRefresh.Click
        FillTableAdapter()
    End Sub
    Private Sub tsbUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbUpdate.Click
        Me.UpdateData()
    End Sub
#End Region

#Region "Audio"
    Public Overridable Function blnPlayAudio() As Boolean
        Return False
    End Function
#End Region

#Region "InThread"
    Public Overridable Sub TimerInMainThread()
        Debug.Print("Timer fired")
    End Sub
    Public Overridable Sub TimerOutSideThread()

    End Sub
#End Region

#Region "KeyBord"
    '20120811 
    Public Overridable Function frmStandard_EnterKey() As Boolean
        Return False
    End Function
    Public Overridable Function frmStandard_F10Key() As Boolean
        FillTableAdapter()
        Return True
    End Function

    Public Overridable Function frmStandard_F9Key() As Boolean
        Dim blnRO As Boolean = True
        If statics.blnCheckLevel(MainForm.SECURITY, blnRO) Then
            MainDefs.MainForm.CloseIfExists(MainDefs.MainForm, MainForm.SECURITY)

            '20130131 strForm instead of strSecurityName ub frmManage constructor so that it works on MasterData as well as forms.
            MainDefs.MainForm.DisplayAForm(MainDefs.MainForm, New frmManage(MainForm.SECURITY, MainDefs, strForm), MainForm.SECURITY, MainForm.SECURITY)
        End If
        Return True
    End Function

    'only works because Me.KeyPreview = True is set in New'
    Private Sub frmStandard_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = frmStandard_EnterKey()
        Else
            If e.KeyCode = Keys.F9 Then
                e.Handled = frmStandard_F9Key()
            End If
            If e.KeyCode = Keys.F10 Then
                e.Handled = frmStandard_F10Key()
            End If
        End If
    End Sub

    'this demonstrates how to kidnap keys before they reach the underlying controls.
    'Public Overridable Sub Form1_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles Me.KeyPress
    '    If e.KeyChar >= ChrW(48) And e.KeyChar <= ChrW(57) Then
    '        MessageBox.Show(("Form.KeyPress: '" + _
    '            e.KeyChar.ToString() + "' pressed."))

    '        Select Case e.KeyChar
    '            Case ChrW(49), ChrW(52), ChrW(55)
    '                MessageBox.Show(("Form.KeyPress: '" + _
    '                    e.KeyChar.ToString() + "' consumed."))
    '                e.Handled = True
    '        End Select
    '    End If
    'End Sub
#End Region

#Region "ControlManagement"
    'set controls to RO if necessary, set tooltip to table.field if bound. 
    Private Sub ModifyAssociatedLabel(ByVal Controls As System.Windows.Forms.Control.ControlCollection, ByVal LeafComponents As System.ComponentModel.IContainer, _
    ByVal ctl As Control, _
    ByVal strTable As String, ByVal strField As String, ByVal strHeader As String, ByVal blnVisible As Boolean)

        'look up associated label if the control is bound.
        If strTable.Length > 0 And strField.Length > 0 Then
            'try to find the label associated with a Text, Combo or Check box.
            Dim strLabel As String = ""
            Dim strCtrlName As String = ""
            If ctl.Name.Contains("TextBox") Then
                strLabel = ctl.Name.Replace("TextBox", "Label")
                strCtrlName = ctl.Name.Replace("TextBox", "")
            Else
                If ctl.Name.Contains("ComboBox") Then
                    strLabel = ctl.Name.Replace("ComboBox", "Label")
                    strCtrlName = ctl.Name.Replace("ComboBox", "")
                Else
                    If ctl.Name.Contains("CheckBox") Then
                        strLabel = ctl.Name.Replace("CheckBox", "Label")
                        strCtrlName = ctl.Name.Replace("CheckBox", "")
                    Else
                        If ctl.Name.Contains("ListBox") Then
                            strLabel = ctl.Name.Replace("ListBox", "Label")
                            strCtrlName = ctl.Name.Replace("ListBox", "")
                        End If
                    End If
                End If
            End If

            If strCtrlName = "" Then
                If ctl.Name.StartsWith("tb") Then
                    strLabel = "l" + ctl.Name.Substring(2)
                Else
                    If ctl.Name.StartsWith("cb") Then
                        strLabel = "l" + ctl.Name.Substring(2)
                    Else
                        If ctl.Name.StartsWith("lb") Then
                            strLabel = "l" + ctl.Name.Substring(2)
                        End If
                    End If
                End If
            End If
            If strCtrlName = "" Then strCtrlName = strLabel


            If strLabel.Length > 0 Then
                For Each Labelctl As Control In Controls
                    If Labelctl.Name.StartsWith(strLabel) Or Labelctl.Name.Contains(strCtrlName) Then
                        Dim lbl As Label = TryCast(Labelctl, Label)
                        If Not lbl Is Nothing Then

                            'Labelctl.Anchor = AnchorStyles.Right
                            Dim iW = Labelctl.Width

                            'default text is the field name.
                            'Labelctl.Text = strField

                            'tooltip on the label is table.field
                            Dim ToolT As ToolTip = New System.Windows.Forms.ToolTip(LeafComponents)
                            ToolT.SetToolTip(Labelctl, strTable + "." + strField)

                            'then look up the users text.
                            'statics.get_v_tble_column_header(strTable.ToUpper, strField.ToUpper, Labelctl.Text)
                            Labelctl.Text = strHeader
                            Labelctl.Visible = blnVisible
                            Labelctl.Location = New System.Drawing.Point(Labelctl.Location.X + iW - Labelctl.Width, Labelctl.Location.Y)

                            'Exit For
                        End If
                    End If
                Next
            End If
        End If
    End Sub
    Private Sub SetWidth(ByVal ctl As Control, ByVal strTable As String, ByVal strField As String, ByVal iWidth As Integer)
        'only set width if the control is bound.
        If strTable.Length > 0 And strField.Length > 0 Then
            ctl.Width = iWidth
        End If
    End Sub

#End Region

#Region "Groupboxes and Translate"
    Private Sub DisableButtons()

        For Each ctl As Control In Me.Controls
            Dim btn As Button = TryCast(ctl, Button)
            If Not btn Is Nothing Then
                If blnRO = True Then
                    btn.Enabled = False
                End If
            End If
        Next

    End Sub

    Private Sub TranslateControl(ctl As Control)
        Dim iL As Integer = ctl.Text.Length

        'get the new or translated name.
        ctl.Text = statics.get_txt_header(ctl.Text, "Control text", strSecurityName)

        'Try and resize the button if the button name is longer than the original.
        If iL < ctl.Text.Length Then
            If ctl.Size.Width < ctl.Text.Length * 7 Then
                ctl.Size = New System.Drawing.Size(ctl.Size.Width + (ctl.Text.Length - iL) * 7, ctl.Size.Height)
            End If
        End If

    End Sub
    Private Sub TranslateGroupBoxCtrls(ByVal LeafComponents As System.ComponentModel.IContainer, ByVal ctl As Control)
        Try

            'Here is a recursive call for groupboxes in groupboxes.
            Dim gb As GroupBox = TryCast(ctl, GroupBox)
            If Not gb Is Nothing Then
                TranslateGroupBox(LeafComponents, gb)
                Return
            End If

            Dim btn As Button = TryCast(ctl, Button)
            If Not btn Is Nothing Then
                TranslateControl(ctl)
            End If
            Dim lab As Label = TryCast(ctl, Label)
            If Not lab Is Nothing Then
                TranslateControl(ctl)
            End If

            'translate the groupbox name after doing the above.
            If gb.Text.Length > 0 Then
                gb.Text = statics.get_txt_header(gb.Text, "groupbox name", strSecurityName)

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub TranslateGroupBox(ByVal LeafComponents As System.ComponentModel.IContainer _
                                  , ByVal gb As Control)


        'then do the controls in the groupbox.
        For Each gbctl As Control In gb.Controls
            Dim gblb As GroupBox = TryCast(gbctl, GroupBox)
            If Not gblb Is Nothing Then
                TranslateGroupBox(LeafComponents, gbctl)
            End If

            Dim btn As Button = TryCast(gbctl, Button)
            If Not btn Is Nothing Then
                TranslateControl(btn)
            End If
            Dim lab As Label = TryCast(gbctl, Label)
            If Not lab Is Nothing Then
                TranslateControl(lab)
            End If
        Next

    End Sub

    Public Sub TranslateAllText(ByVal LeafComponents As System.ComponentModel.IContainer)
        Try
            For Each ctl As Control In Me.Controls
                Dim gb As GroupBox = TryCast(ctl, GroupBox)
                If Not gb Is Nothing Then
                    TranslateGroupBox(LeafComponents, gb)
                Else
                    Dim flp As FlowLayoutPanel = TryCast(ctl, FlowLayoutPanel)
                    If Not flp Is Nothing Then
                        TranslateGroupBox(LeafComponents, flp)
                    Else
                        Dim btn As Button = TryCast(ctl, Button)
                        If Not btn Is Nothing Then
                            TranslateControl(ctl)
                        Else
                            Dim lab As Label = TryCast(ctl, Label)
                            If Not lab Is Nothing Then
                                TranslateControl(ctl)
                            End If
                        End If
                    End If
                End If
            Next
        Catch ex As Exception

        End Try
    End Sub


    ''' <summary>
    ''' Get visibility etc from the Utilities form_tble_column table for the field if the control is bound.
    ''' Cast the control to find out what sort of control it is. Adjust visibility and then try to find the assoicated label adjusting it as well.
    ''' </summary>
    ''' <param name="Controls"></param>
    ''' <param name="ctl"></param>
    ''' <param name="LeafComponents"></param>
    ''' <remarks>Controls is explicitly the groupbox Controls declaration. The Form Groupbox is different but casts to the groupbox control.</remarks>
    Private Sub CreateLabelAndTextBoxTooltips(ByVal Controls As System.Windows.Forms.Control.ControlCollection, _
        ByVal ctl As Control, ByVal LeafComponents As System.ComponentModel.IContainer, ByVal blnCtlRO As Boolean)
        Try
            Dim strTable As String = ""
            Dim strField As String = ""
            Dim strHeader As String = ""
            Dim strFormat As String = "" 'not supported here
            Dim iwidth As Integer = 0
            Dim blnVisible As Boolean = True
            Dim blnPrnt As Boolean = True 'not supported here
            Dim blnBold As Boolean = False 'not supported here
            Dim iSequence As Integer 'not supported here

            If ctl.DataBindings.Count > 0 Then
                If Not ctl.DataBindings(0).DataSource.datamember Is Nothing And Not ctl.DataBindings(0).BindingMemberInfo.BindingField Is Nothing Then
                    strTable = ctl.DataBindings(0).DataSource.datamember
                    strField = ctl.DataBindings(0).BindingMemberInfo.BindingField
                    statics.get_v_form_tble_column(strSecurityName, strTable, strField, _
                        strHeader, strFormat, iwidth, blnVisible, blnPrnt, blnBold, iSequence)
                End If
            End If

            Dim comboB As ComboBox = TryCast(ctl, ComboBox)
            If Not comboB Is Nothing Then
                SetWidth(ctl, strTable, strField, iwidth)
                ModifyAssociatedLabel(Controls, LeafComponents, ctl, strTable, strField, strHeader, blnVisible)
                If comboB.Enabled = True Then
                    comboB.Enabled = Not blnCtlRO
                End If
                comboB.Visible = blnVisible
                Return
            End If

            Dim textB As TextBox = TryCast(ctl, TextBox)
            If Not textB Is Nothing Then
                SetWidth(ctl, strTable, strField, iwidth)
                ModifyAssociatedLabel(Controls, LeafComponents, ctl, strTable, strField, strHeader, blnVisible)
                If textB.ReadOnly = False Then 'only make RO or not RO is not set to RO in designer.
                    textB.ReadOnly = blnCtlRO
                End If
                textB.Visible = blnVisible
                Return
            End If

            Dim checkB As CheckBox = TryCast(ctl, CheckBox)
            If Not checkB Is Nothing Then
                SetWidth(ctl, strTable, strField, iwidth)
                ModifyAssociatedLabel(Controls, LeafComponents, ctl, strTable, strField, strHeader, blnVisible)
                If checkB.Enabled = True Then
                    checkB.Enabled = Not blnCtlRO
                End If
                checkB.Visible = blnVisible
                Return
            End If

            Dim ListB As ListBox = TryCast(ctl, ListBox)
            If Not ListB Is Nothing Then
                SetWidth(ctl, strTable, strField, iwidth)
                ModifyAssociatedLabel(Controls, LeafComponents, ctl, strTable, strField, strHeader, blnVisible)
                If ListB.Enabled = True Then
                    ListB.Enabled = Not blnCtlRO
                End If
                ListB.Visible = blnVisible
                Return
            End If

            Dim NUpDown As NumericUpDown = TryCast(ctl, NumericUpDown)
            If Not NUpDown Is Nothing Then
                SetWidth(ctl, strTable, strField, iwidth)
                ModifyAssociatedLabel(Controls, LeafComponents, ctl, strTable, strField, strHeader, blnVisible)
                If NUpDown.ReadOnly = False Then
                    NUpDown.ReadOnly = blnCtlRO
                End If
                NUpDown.Visible = blnVisible
                Return
            End If

            'Here is a recursive call for groupboxes in groupboxes.
            Dim gb As GroupBox = TryCast(ctl, GroupBox)
            If Not gb Is Nothing Then
                ProcessGroupBox(LeafComponents, gb, blnCtlRO)
                Return
            End If

            Dim btn As Button = TryCast(ctl, Button)
            If Not btn Is Nothing Then
                'Dim strButtonName As String = Me.strSecurityName
                'If Controls.Owner.Text <> Me.Text Then
                '    strButtonName = strButtonName + "." + Controls.Owner.Text
                'End If
                'strButtonName = strButtonName + "." + btn.Text
                'store the name of the button the first time around.
                'statics.put_v_form(strButtonName, False, False)
                'Dim blnBRO As Boolean
                'If statics.blnCheckLevel(strButtonName, blnBRO) = False Then

                If blnCtlRO = True Then
                    btn.Enabled = False
                End If

                Dim iL As Integer = btn.Text.Length

                'set tooltip text to the original name.
                Dim ToolT As ToolTip = New System.Windows.Forms.ToolTip(LeafComponents)
                ToolT.SetToolTip(btn, btn.Text)

                'get the new or translated name.
                btn.Text = statics.get_txt_header(btn.Text, "button text", strSecurityName)

                'Try and resize the button if the button name is longer than the original.
                If iL < btn.Text.Length Then
                    If btn.Size.Width < btn.Text.Length * 7 Then
                        btn.Size = New System.Drawing.Size(btn.Size.Width + (btn.Text.Length - iL) * 7, btn.Size.Height)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub ProcessGroupBox(ByVal LeafComponents As System.ComponentModel.IContainer, ByVal gb As GroupBox, ByVal blnCtlRO As Boolean)
        Dim ToolT As ToolTip = New System.Windows.Forms.ToolTip(LeafComponents)
        ToolT.SetToolTip(gb, gb.Text)

        gb.ForeColor = statics.GetTextColor(gb.BackColor)

        'then do the controls in the groupbox.
        For Each gbctl As Control In gb.Controls
            Dim gblb As Label = TryCast(gbctl, Label)
            'If Not gblb Is Nothing Then
            '    Dim gbToolT As ToolTip = New System.Windows.Forms.ToolTip(LeafComponents)
            '    ToolT.SetToolTip(gblb, gblb.Text)
            'Else
            CreateLabelAndTextBoxTooltips(gb.Controls, gbctl, LeafComponents, blnCtlRO)
            'End If
        Next

        'translate the groupbox name after doing the above.
        If gb.Text.Length > 0 Then
            gb.Text = statics.get_txt_header(gb.Text, "groupbox name", strSecurityName)

        End If

    End Sub

    ''' <summary>
    ''' Couple tooltips to Groupbox and labels so that a user can see the name of the field.
    ''' Translate groupbox name and label text.
    ''' Set the width of text and combo boxes.
    ''' Set the controls in a groupbox to read only if checked in m_form_grp_groupbox.
    ''' </summary>
    ''' <param name="LeafComponents"></param>
    ''' <remarks>Add the dialog name and the dialog name, groupbox to the list of form_grp_groupbox. </remarks>
    Protected Sub CreateTooltips(ByVal LeafComponents As System.ComponentModel.IContainer)
        Try
            If Not LeafComponents Is Nothing Then
                For Each ctl As Control In Controls
                    Dim gb As GroupBox = TryCast(ctl, GroupBox)
                    If Not gb Is Nothing Then
                        Dim blnGBRO As Boolean = blnRO

                        'store the name of the form.group the first time it is opened with ro = false.
                        statics.put_v_form_groupbox(strSecurityName, gb.Text)
                        If gb.Text.Length > 0 And blnRO = False Then

                            'then check the level of the form.group. If the user may not use the controls make RO true.
                            blnGBRO = statics.blnCheckGroupBoxLevel(strSecurityName, gb.Text)
                        End If
                        ProcessGroupBox(LeafComponents, gb, blnGBRO)

                        'translate the groupbox name after doing the above.
                        If gb.Text.Length > 0 Then
                            gb.Text = statics.get_txt_header(gb.Text, "groupbox name", strSecurityName)
                        End If
                    Else
                        CreateLabelAndTextBoxTooltips(Controls, ctl, LeafComponents, blnRO)
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Is needed to adjust the group box layout and field visibility. Trying to do this on the complete form 
    'can cause problems in a groupbox which should not be checked for example when 
    'a datagrid field is not visible in the grid but should be visible in a textbox in the same groupbox.
    Public Sub CreateTooltips(gb As GroupBox)
        Try
            Dim LeafComponents As System.ComponentModel.IContainer = Me.components
            If Not LeafComponents Is Nothing Then
                If Not gb Is Nothing Then
                    Dim blnGBRO As Boolean = blnRO


                    'store the name of the form.group the first time it is opened with ro = false.
                    statics.put_v_form_groupbox(strSecurityName, gb.Text)
                    If gb.Text.Length > 0 And blnRO = False Then

                        'then check the level of the form.group. If the user may not use the controls make RO true.
                        blnGBRO = statics.blnCheckGroupBoxLevel(strSecurityName, gb.Text)
                    End If
                    ProcessGroupBox(LeafComponents, gb, blnGBRO)

                    'translate the groupbox name after doing the above.
                    If gb.Text.Length > 0 Then
                        gb.Text = statics.get_txt_header(gb.Text, "groupbox name", strSecurityName)
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

End Class