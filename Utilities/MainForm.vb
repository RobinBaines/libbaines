'------------------------------------------------
'Name: Module MainForm.vb.
'Function: The class which is inherited when creating a new mainform.
'Copyright Robin Baines 2006. All rights reserved.
'------------------------------------------------
'20090818 RPB added code in BroadcastFilter to support filtering between all open forms.
'------------------------------------------------
'20091117 RefreshStatusStrip() set to Overridable for use in TPIStatus.
'20091209 RPB added AddToolStripRange with list of ToolStripItems.
'20120510 RPB changed the Security related buttons to ToolStripItems and put in a drop down on the 
'main menu (ToolStrip1), see tsbMaintain.
'When doing the same in a derived main form call CreateMainMenuDropDownEntry
'from Public Overrides Sub CheckFormVisibility(). This ensures that translation etc happens after the 
'forms with translation text have been read from the database.
'Moved the tsl labels with version number to the Status bar at the bottom of the form.
'20120521 CreateMainMenuDropDownEntry modified.
'20120529 Changed tsb from  ToolStripButton to ToolStripItem for more generic switching of font when an item is selected.

'When DELETING Forms from Security form remember that the Forms only get created on initialisation so 
'it is necessary to restart the application to re-create.
' 20150815 added parameter lsemaphore when distributing event because semaphore has altered.
'20180222 According to MS, there is a known issue with the DataGrid control in certain situations, and the way to avoid this issue is to disable Tooltips for your DataGrid controls.
'But also Adding Application.EnableVisualStyles() in main() before any forms are instantiated solves the problem.
'20180225 set up exception handling to catch any exception.
'20200219 frmManage: Do not allow dialog DialogSelectForms to be shown if the Security form is Read only.
'20200123 dgColumns: In some circumstances functions may be called when the form is being disposed. So check dgSortedListOfColumns before disposing the columns.
'20200222 dgColumns: So check gbForFiltersGroupBox before disposing the columns.
'20211112 Added try/catch with msgbox during initialisation calls Init() and ModuleStatics.InitConnection().
'------------------------------------------------
Imports System.Windows.Forms
Imports System.Windows
Imports System.Drawing
Imports System.Threading
Imports System
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Reflection
Imports System.Diagnostics
Public Class MainForm

    Public MainDefs As MainDefinitions
    Protected WithEvents tsbBtnResetFilter As System.Windows.Forms.ToolStripButton
    Protected WithEvents tslUser As System.Windows.Forms.ToolStripLabel
    Protected WithEvents tslVersion As System.Windows.Forms.ToolStripLabel
    Protected WithEvents tslTimeToGo As System.Windows.Forms.ToolStripLabel
    Protected WithEvents tslTestDatabase As System.Windows.Forms.ToolStripLabel
    Protected WithEvents tslRam As System.Windows.Forms.ToolStripLabel
    Friend WithEvents tsmProperties As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmHelpText As System.Windows.Forms.ToolStripMenuItem


    Dim blnStartThread As Boolean = False
    Dim strNameOfForm As String
    Protected strAudioFile As String = ""
    Private _SQLStatus As statics.StatusValues
    Const STRRESETNAME As String = "Reset"

    Friend WithEvents tsbMaintain As System.Windows.Forms.ToolStripMenuItem
    Const MAINTAIN As String = "Maintain"
    Public Const SECURITY As String = "Security"
    Const USERLOG As String = "User Log"
    Const OPTIONS As String = "Options"
    Const APPLOG As String = "App Log"

    'Protected WithEvents tsbMeta As System.Windows.Forms.ToolStripButton
    Const MAINMENU_META_VIEWS As String = "View Browser"

    'Protected WithEvents tsbMetaProc As System.Windows.Forms.ToolStripButton
    Const MAINMENU_META_Procs As String = "Routines Browser"

    Const MENU_ADHOC_VIEWS As String = "Adhoc Views"

    Public SQLParser As clParseSQL

#Region "Properties"
    Public Property SQLStatus() As statics.StatusValues
        Get
            Return _SQLStatus
        End Get
        Set(ByVal value As statics.StatusValues)
            _SQLStatus = value
        End Set
    End Property

    Public ReadOnly Property iTimerTick() As String
        Get
            Return "TimerTick"
        End Get
    End Property

    Public ReadOnly Property iShortTimerTick() As String
        Get
            Return "ShortTimerTick"
        End Get
    End Property

    Private ReadOnly Property strSuperShortTimerTick() As String
        Get
            Return "SuperShortTimerTick"
        End Get
    End Property
#End Region

#Region "New"
    Public Sub New()
        MyBase.New()
        Try
            InitializeComponent()
            NewInit()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    ''' <summary>
    ''' 20100604 allow explicit starting of the thread.
    ''' </summary>
    ''' <param name="_blnStartThread"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal _blnStartThread As Boolean)
        MyBase.New()
        Try
            InitializeComponent()
            NewInit()
            blnStartThread = _blnStartThread
            SQLStatus = statics.StatusValues.OK

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub DomainExceptionHandler(ByVal sender As Object, _
ByVal e As UnhandledExceptionEventArgs)
        Dim EX As Exception
        EX = e.ExceptionObject
        MsgBox(EX.StackTrace)
    End Sub

    Private Sub ApplicationThreadExceptionHandler(ByVal sender As Object, _
      ByVal e As Threading.ThreadExceptionEventArgs)
        MsgBox(e.Exception.StackTrace)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overridable Sub NewInit()

        '20180222 According to MS, there is a known issue with the DataGrid control in certain situations, and the way to avoid this issue is to disable Tooltips for your DataGrid controls.
        'But also Adding Application.EnableVisualStyles() in main() before any forms are instantiated solves the problem.
        Application.EnableVisualStyles()

        '20180225 set up exception handling to catch any exception.
        Dim currentDomain As AppDomain = AppDomain.CurrentDomain

        ' Define a handler for unhandled exceptions.
        AddHandler currentDomain.UnhandledException, AddressOf DomainExceptionHandler

        ' Define a handler for unhandled exceptions for threads behind forms.
        AddHandler Application.ThreadException, AddressOf ApplicationThreadExceptionHandler

        statics.Init(Me)
        'MsgBox("Before TheMenustrip()")
        TheMenustrip()
        'MsgBox("Before TheToolstrip()")
        TheToolstrip()
        'MsgBox("Before MainDefinitions()")
        MainDefs = New MainDefinitions(Me)

        'store the name of the form so that it is available after Text has been translated.
        strNameOfForm = strGetMainFormName()

    End Sub

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        MyBase.OnClosing(e)
        statics.Closing()
    End Sub

    Protected Overridable Function strGetMainFormName() As String
        Return Application.ProductName
        'Return (Me.Text)
    End Function

    Public Overridable Sub TheMenustrip()

        'Properties to switch to/from test database.
        tsmProperties = CreateTsmForMenuStrip("tsmProperties", "Properties", True)
        tsmHelpText = CreateTsmForMenuStrip("tsmHelpText", "Help Text", True)
        Me.MenuStrip.ShowItemToolTips = True
    End Sub

    Public Overridable Sub TheToolstrip()

        Me.tslUser = New System.Windows.Forms.ToolStripLabel
        Me.tslVersion = New System.Windows.Forms.ToolStripLabel
        Me.tslTimeToGo = New System.Windows.Forms.ToolStripLabel
        tslTestDatabase = New System.Windows.Forms.ToolStripLabel
        tslRam = New System.Windows.Forms.ToolStripLabel

        'Create and add to Toolstrip1 immediately
        tsbBtnResetFilter = CreateTsb("tsbBtnResetFilter", STRRESETNAME, True, True)

        Me.tslUser.BackColor = System.Drawing.SystemColors.Control
        Me.tslUser.Font = New System.Drawing.Font("Tahoma", 8.400001!)
        Me.tslUser.Name = "tslUser"
        Me.tslUser.Size = New System.Drawing.Size(43, 25)
        Me.tslUser.Text = ""
        tslUser.BackColor = StatusStrip.BackColor

        Me.tslVersion.Font = New System.Drawing.Font("Tahoma", 8.400001!)
        Me.tslVersion.Name = "tslVersion"
        Me.tslVersion.Size = New System.Drawing.Size(55, 25)
        Me.tslVersion.Text = ""
        tslVersion.BackColor = Color.Silver

        Me.tslTestDatabase.Font = New System.Drawing.Font("Tahoma", 8.400001!)
        Me.tslTestDatabase.Name = "tslTestDatabase"
        Me.tslTestDatabase.Size = New System.Drawing.Size(55, 25)
        Me.tslTestDatabase.Text = ""
        tslTestDatabase.BackColor = StatusStrip.BackColor

        Me.tslRam.Font = New System.Drawing.Font("Tahoma", 8.400001!)
        Me.tslRam.Name = "tslRam"
        Me.tslRam.Size = New System.Drawing.Size(137, 25)
        Me.tslRam.Text = ""
        tslRam.BackColor = Color.Silver

        Me.tslTimeToGo.Font = New System.Drawing.Font("Tahoma", 8.400001!)
        Me.tslTimeToGo.Name = "tsbTimeToGo"
        Me.tslTimeToGo.Size = New System.Drawing.Size(137, 25)
        Me.tslTimeToGo.Text = ""
        tslTimeToGo.BackColor = StatusStrip.BackColor

    End Sub

#End Region

#Region "AdjustButtons"
    ''' <summary>
    ''' Allow user to add an item to the MenuStrip
    ''' </summary>
    ''' <param name="tsItems"></param>
    ''' <remarks></remarks>
    Protected Overridable Sub TheMenustrip(ByVal tsItems As ToolStripItem())
        Me.MenuStrip.Items.AddRange(tsItems)
    End Sub

    ''' <summary>
    ''' Allow application to switch off main menu items. Do after check on visibility as that wil probably make the visible.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub SwitchOffFile()
        Me.FileMenu.Visible = False
    End Sub

    Protected Sub SwitchOffEdit()
        Me.EditMenu.Visible = False
    End Sub

    Protected Sub SwitchOffTools()
        Me.ToolsMenu.Visible = False
    End Sub

    Protected Sub SwitchOffHelp()
        Me.HelpMenu.Visible = False
    End Sub

    Protected Sub SwitchOffView()
        Me.ViewMenu.Visible = False
    End Sub

    Protected Sub SwitchOffWindows()
        Me.WindowsMenu.Visible = False
    End Sub

    Protected Sub SwitchOffStatus()
        'The status strip at the bottom of the form.
        Me.StatusStrip.Visible = False
    End Sub
    Protected Sub SwitchOffProperties()
        Me.tsmProperties.Visible = False
    End Sub

    Protected Sub SwitchOffHelpText()
        Me.tsmHelpText.Visible = False
    End Sub

    Protected Sub SwitchOffReset()
        tsbBtnResetFilter.Visible = False
    End Sub
    Protected Sub SwitchOffMaintain()
        tsbMaintain.Visible = False
    End Sub

    ''' <summary>
    ''' 20091209 RPB added AddToolStripRange with list of ToolStripItems. This is to allow Toolstrip items after custom and generic ones.
    ''' </summary>
    ''' <param name="tsItems"></param>
    ''' <remarks></remarks>
    Protected Overridable Sub AddToolStripRange(ByVal tsItems As ToolStripItem())
        Me.ToolStrip1.Items.AddRange(tsItems)
    End Sub

    Protected Overridable Sub AddStatusStripRange(ByVal tsItems As ToolStripItem())
        Me.StatusStrip.Items.AddRange(tsItems)
    End Sub

    ''' <summary>
    ''' Add generic toolstrip items.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub AddStatusStripItems()

        'Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslUser, Me.tslVersion}) ', tslTestDatabase, Me.tslTimeToGo})
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslUser, Me.tslVersion}) ', tslTestDatabase, Me.tslTimeToGo})
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {tslTestDatabase, Me.tslTimeToGo, tslRam})

    End Sub

    Private Sub SwitchOffAllMenuButtons()

        Dim tsb As ToolStripButton

        'blocked or not authorized so make all menu buttons invisible.
        For Each ctl As ToolStripItem In Me.ToolStrip1.Items
            tsb = TryCast(ctl, ToolStripButton)
            If Not tsb Is Nothing Then
                tsb.Visible = False
            End If
        Next
        Dim tsm As ToolStripMenuItem
        For Each ctl As ToolStripItem In Me.MenuStrip.Items
            tsm = TryCast(ctl, ToolStripMenuItem)
            If Not tsm Is Nothing Then
                tsm.Visible = False
            End If
        Next

    End Sub
    Private Function GetColour(iColorSwitch As Integer) As Color
        If iColorSwitch = 0 Then
            Return Color.AliceBlue
        Else
            If iColorSwitch = 1 Then
                Return Color.Lavender
            Else

                If iColorSwitch = 2 Then
                    Return Color.Khaki
                Else

                    If iColorSwitch = 3 Then
                        Return Color.LightPink

                    Else
                        If iColorSwitch = 4 Then
                            Return Color.Moccasin
                        End If
                    End If
                End If
            End If
        End If
        Return Color.White
    End Function

    Private Sub AdjustColorAndText(ctl As ToolStripItem, ByRef iColorswitch As Integer)
        Dim strSaveToolTip As String = ctl.ToolTipText
        ctl.Text = statics.get_txt_header(ctl.ToolTipText)
        ctl.ToolTipText = strSaveToolTip
        If iColorswitch > -1 Then
            ctl.BackColor = GetColour(iColorswitch)
            iColorswitch += 1
            If iColorswitch >= 4 Then
                iColorswitch = 0
            End If
        End If
    End Sub

    Private Function IsDefaultMenuOption(tsm As ToolStripMenuItem) As Boolean
        Select Case tsm.Name
            Case "FileMenu"
                Return True
            Case "EditMenu"
                Return True
            Case "ToolsMenu"
                Return True
            Case "WindowsMenu"
                Return True
            Case "ViewMenu"
                Return True
            Case "HelpMenu"
                Return True
        End Select
        Return False

    End Function


    ''' <summary>
    ''' Check the form visibility using the security tables.
    ''' </summary>
    ''' <remarks>This gets called when switching to/from test database so the ToolTip texts are set the first
    ''' time and are used to look up the translation.</remarks>
    '''    'The ToolStripItem Available property is different from the Visible property in that Available indicates 
    '''    'whether the ToolStripItem is shown, while Visible indicates whether the ToolStripItem and 
    '''    'its parent are shown. 
    '''    'Setting either Available or Visible to true or false sets the other property to true or false.

    Public Overridable Sub CheckFormVisibility()
        Dim iColorSwitch As Integer = 0
        Try


            If statics.Blocked = False Then
                If statics.IsAuthorized = True Then

                    '20120912 moved this here so it is not added to the menu if blocked or not authorized.
                    'Create this here to force it to be the last entry.
                    If tsbMaintain Is Nothing Then
                        tsbMaintain = CreateMainMenuDropDownEntry(Nothing, MAINTAIN)
                        CreateMainMenuDropDownEntry(tsbMaintain, SECURITY)
                        CreateMainMenuDropDownEntry(tsbMaintain, USERLOG)
                        CreateMainMenuDropDownEntry(tsbMaintain, OPTIONS)
                        CreateMainMenuDropDownEntry(tsbMaintain, MENU_ADHOC_VIEWS)
                        CreateMainMenuDropDownEntry(tsbMaintain, MAINMENU_META_VIEWS)
                        CreateMainMenuDropDownEntry(tsbMaintain, MAINMENU_META_Procs)

                        If statics.blnSupportsAppLog = True Then
                            CreateMainMenuDropDownEntry(tsbMaintain, APPLOG)
                        End If
                    End If

                    For Each ctl As ToolStripItem In Me.ToolStrip1.Items
                        Dim tsb As ToolStripButton = Nothing
                        tsb = TryCast(ctl, ToolStripButton)
                        If Not tsb Is Nothing Then

                            'THIS IS ALSO DONE IN CreateTsb but need it here too because not all tsbs are being created
                            'using CreateTsb. 
                            'Lookup language dependent name for the tsb text.
                            'the ToolTip text is set when creating the tsb's.
                            'Initialise the ToolTipText the first time.
                            If ctl.ToolTipText Is Nothing Then
                                ctl.ToolTipText = tsb.Text
                            End If

                            '20120712 Do not add Reset to the database.
                            Dim blnRet As Boolean = True
                            If Not tsb Is Nothing Then
                                If tsb.Text <> STRRESETNAME Then
                                    statics.put_v_form(tsb.ToolTipText, False, False)   'tsb.Text, False, False)
                                    blnRet = statics.blnCheckLevel(tsb.ToolTipText)
                                    ctl.Visible = blnRet
                                End If
                            End If
                            If blnRet = True Then
                                AdjustColorAndText(ctl, iColorSwitch)
                            End If
                        Else
                            'Get a toolstripmenu item.
                            Dim tsmMainMenu As ToolStripMenuItem = Nothing
                            tsmMainMenu = TryCast(ctl, ToolStripMenuItem)
                            If Not tsmMainMenu Is Nothing Then
                                If tsmMainMenu.HasDropDownItems = True Then

                                    'If this one has children then set it on if at least one child is visible.
                                    Dim blnParentIsVis As Boolean = False
                                    Dim iChildColorSwitch As Integer = 0
                                    For Each ctl2 As ToolStripItem In tsmMainMenu.DropDownItems
                                        Dim tsm2 As ToolStripMenuItem = TryCast(ctl2, ToolStripMenuItem)
                                        If Not tsm2 Is Nothing Then
                                            Dim blnChildIsVis As Boolean = statics.blnCheckLevel(tsm2.ToolTipText)
                                            If blnChildIsVis = True Then
                                                blnParentIsVis = blnChildIsVis
                                            End If
                                            tsm2.Visible = blnChildIsVis
                                            AdjustColorAndText(tsm2, iChildColorSwitch)
                                        End If
                                    Next
                                    tsmMainMenu.Visible = blnParentIsVis
                                    If blnParentIsVis = True Then
                                        AdjustColorAndText(tsmMainMenu, iColorSwitch)
                                    End If
                                Else

                                    '20130913 If this one has no child items then set it off.
                                    Dim blnRet As Boolean = False
                                    'If tsb.Text <> STRRESETNAME Then
                                    '    blnRet = statics.blnCheckLevel(tsb.ToolTipText)
                                    '    ctl.Visible = blnRet
                                    'End If
                                    ctl.Visible = False
                                    'If blnRet = True Then
                                    '    AdjustColorAndText(ctl, iColorSwitch)
                                    'End If
                                End If
                            End If

                        End If
                    Next

                    'And now the menu strip.
                    For Each ctl As ToolStripItem In Me.MenuStrip.Items
                        Dim tsm As ToolStripMenuItem = Nothing
                        tsm = TryCast(ctl, ToolStripMenuItem)
                        If Not tsm Is Nothing Then

                            'don't bother with the default menu options.
                            'these depend on the application and are switched off from app main if necessary.
                            If Not IsDefaultMenuOption(tsm) Then
                                If Not tsm.ToolTipText Is Nothing Then
                                    tsm.Visible = statics.blnCheckLevel(tsm.ToolTipText)
                                Else
                                    If tsm.Name.Length > 3 Then
                                        tsm.Visible = statics.blnCheckLevel(tsm.Name.Substring(3))
                                    End If
                                End If

                                'lookup language dependent name for the tsb text.
                                'Initialise the ToolTipText the first time.
                                If tsm.ToolTipText Is Nothing Then
                                    tsm.ToolTipText = tsm.Text
                                End If
                                If tsm.Visible = True And Not tsm.Text.StartsWith("&") Then
                                    AdjustColorAndText(tsm, -1)
                                End If
                            Else
                            End If
                        End If
                    Next
                Else
                    SwitchOffAllMenuButtons()

                    'prompt user so they know what is going on.
                    MsgBox(statics.get_txt_header("UserNotAuthorized"))
                End If
            Else
                SwitchOffAllMenuButtons()

                'prompt user so they know what is going on.
                MsgBox(statics.get_txt_header("UserBlocked"))
            End If
        Catch ex As Exception
            MsgBox("problem " + ex.Message)
        End Try
    End Sub
#End Region

#Region "Load"
    ''' <summary>
    ''' Used when switching to/from test database.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CloseAllForms()
        Dim frms() As Form
        Dim f As Form
        Dim fRet As Form = Nothing
        frms = Me.MdiChildren()
        For Each f In frms
            Try

                'Do not close this form. But why is name = Column Design?
                'If f.Name <> "frmColumns" And f.Name <> "Column Design" Then
                f.Close()
                f.Dispose()
                f = Nothing
                'End If
            Catch ex As Exception
            End Try
        Next
    End Sub

    Protected Overridable Function strGetAppDetails() As String
        Return Application.ProductVersion & " of " & Application.ProductName & " by " & Application.CompanyName
    End Function

    Protected Overridable Sub CreateDropDownMenus()

    End Sub

    Protected Sub UpdateRamState()
        Dim x As Process = Process.GetCurrentProcess()

        'Megabyte (MB): 1024KB equals one megabyte (MB), Gigabyte (GB): There are 1024MB in one gigabyte. Hard drive manufacturers have long eschewed this system in favour of rounding down to make things easier (and also provide less storage space). This means that 1000 bytes = 1 kilobyte and 1000 kilobytes = 1MB.
        Dim dWS As Decimal = x.WorkingSet64
        Dim dPM As Decimal = x.PagedMemorySize64
        tslRam.Text = "Mem Usage: " & (dWS / 1000000).ToString("N0") & " M " '& "Paged Memory: " & (dPM / 1000000).ToString("N0") & " M "
    End Sub

    ''' <summary>
    ''' Initialise all items.
    ''' </summary>
    ''' <remarks>Is called when program starts and when switching to/from test database.</remarks>
    Public Overridable Sub Init()

        Dim strUser As String = ""
        Try
            'This one is run once when connecting to a database. So is run when starting but also when 
            'connecting to/from the test database. It writes into the usr_log and updates the statics data.
            Try
                statics.InitConnection()
            Catch ex As Exception
                'if this occurs with the following message then the connection string is incorrect.
                '"A network-related or instance-specific error occurred while establishing a connection to SQL Server. The server was not found or was not accessible. Verify that the instance name is correct and that SQL Server is configured to allow remote connections. (provider: SQL Network Interfaces, error: 26 - Error Locating Server/Instance Specified)"}	System.Exception
                MsgBox("statics.InitConnection " & ex.Message)
            End Try

            Try


                'Switch off Restart controls.
                btnRestart.Visible = False
                Me.tbAppWillEnd.Visible = False
                Me.tbEndSeconds.Visible = False
                Me.tbWillEndSeconds.Visible = False
                If blnStartThread = True Then StartTimer()

                '20120510
                'see documentation on what this does and the options.
                ' FormMenuStrip.LayoutStyle = ToolStripLayoutStyle.StackWithOverflow
                AddStatusStripItems()

                'lookup language dependent name for the tsb text using the strNameOfForm which is initialised before translation.
                Me.Text = statics.get_txt_header(strNameOfForm)
                Me.btnRestart.Text = statics.get_txt_header(Me.btnRestart.Text)
                Me.tbAppWillEnd.Text = statics.get_txt_header(Me.tbAppWillEnd.Text)
                Me.tslVersion.Text = strGetAppDetails()
                Me.tslUser.Text = statics.get_txt_header("User:") & " " & statics.UserName

                '20100603 dont show connection string.
                Me.tslTimeToGo.Text = ""    'GetConnectionString()

            Catch ex As Exception
                MsgBox("Switch off Restart controls. " & ex.Message)
            End Try

            Try
                UpdateRamState()

                'Create the menus before calling init() so that their visibility can be adjusted.
                CreateDropDownMenus()
                CheckFormVisibility()

                If Me.GetQualityDataBase() = True Then
                    tslTestDatabase.Text = "Database is " + statics.get_txt_header("Test")
                    Dim fon As New Font(tslTestDatabase.Font, FontStyle.Bold)
                    tslTestDatabase.Font = fon
                Else
                    tslTestDatabase.Text = "Database is " + statics.get_txt_header("Live data")
                    Dim fon As New Font(tslTestDatabase.Font, FontStyle.Regular)
                    tslTestDatabase.Font = fon
                End If

                RefreshServerStatusStrip("")
                Me.RefreshDirectoryStatusStrip("")
                strAudioFile = statics.GetParameter("audio file")

                SQLParser = New clParseSQL
                SQLParser.Init()
            Catch ex As Exception
                MsgBox("UpdateRamState. " & ex.Message)
            End Try

        Catch ex As Exception
            'if this occurs with the following message then the connection string is incorrect.
            '"A network-related or instance-specific error occurred while establishing a connection to SQL Server. The server was not found or was not accessible. Verify that the instance name is correct and that SQL Server is configured to allow remote connections. (provider: SQL Network Interfaces, error: 26 - Error Locating Server/Instance Specified)"}	System.Exception
            MsgBox("FATAL ERROR: Problem reading from database. " & ex.Message)
            Application.Exit()
        End Try
    End Sub

    ''' <summary>
    ''' Make font of Button normal instead of bold to indicate that form is no longer open.
    ''' </summary>
    ''' <param name="tsb"></param>
    ''' <remarks></remarks>
    Public Overridable Sub DeSelectedFont(ByVal tsb As ToolStripItem)
        Dim SelectFont As System.Drawing.Font
        If Not tsb Is Nothing Then
            SelectFont = New System.Drawing.Font("Tahoma", 8, FontStyle.Regular)
            tsb.Font = SelectFont
        End If
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Overridable Sub MainForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' MsgBox("After NewInit()")
        Init()
        'MsgBox("After Init()")
    End Sub

    'This was in the utilities and is used by TPINet.
    Public Function GetDataSet(ByVal cn As SqlConnection, ByVal sqlQuery As String) As DataSet

        'Execute a select statement and load into a dataset.
        Dim DataSet As New DataSet
        Dim command As New SqlCommand
        command = cn.CreateCommand()
        command.CommandText = sqlQuery
        Dim dataAdapter As SqlDataAdapter
        dataAdapter = New SqlDataAdapter(command)
        dataAdapter.Fill(DataSet)
        'cn.Close()
        Return DataSet
    End Function
#End Region


#Region "IsThisApplicationInForeGround and Icon Overlay"

    Private Declare Function GetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As IntPtr

    'Check whether this application is in the foreground.
    Protected Friend Function IsThisApplicationInForeGround() As Boolean
        Dim blnRet As Boolean = False

        ' get handle of the foreground window
        Dim ForeGroundHwnd As IntPtr = GetForegroundWindow

        ' get processes, find matching foreground window
        Dim procList() = Process.GetProcesses()
        For i As Integer = 0 To procList.Length - 1
            Try
                Dim MainWinHwnd As IntPtr = procList(i).MainWindowHandle
                If MainWinHwnd = ForeGroundHwnd Then
                    Dim procName As String = procList(i).ProcessName
                    Dim procTitle As String = procList(i).MainWindowTitle
                    ' Show some info about the foreground window
                    Debug.Print("ProcName=" & procName & ", Title=" & procTitle & ", main window handle=" & MainWinHwnd.ToString & ", PID=" & procList(i).Id.ToString)
                    If procTitle = Application.ProductName Then
                        blnRet = True
                    End If
                    Exit For ' found so exit loop
                End If
            Catch ex As Exception
            End Try
        Next
        Return blnRet
    End Function

    'THIS HOW TO USE ICON OVERLAY TO SHOW THAT SOMETHING HAS HAPPENED WITH APP IN THE BACKGROUND
    'AND TO RESET IT IF THE APPLICATION COMES BACK.
    'Imports Microsoft.WindowsAPICodePack.Taskbar
    '20131004 Added Taskbar Overlay to signal change to user if LOP is running but not in the foreground (like Outlook does).
    'Downloaded Windows API Code Pack library to Projects\Windows API Code Pack library. The binaries are already available so added 
    '3 references (see references). But also:
    'When using the Windows API Code Pack library to add Taskbar features to your WinForms application you must add a references to the following WPF DLLs: PresentationCore.dll, PresentationFramework and WindowsBase.dll
    'Then added Red ico to resources using Projects.Properties.Resources.Add existing.

    'Private windowsTaskbar As TaskbarManager = TaskbarManager.Instance
    'Public Overrides Sub RefreshAll(ByVal strApp As String, ByVal strTble As String)
    '    MyBase.RefreshAll(strApp, strTble)
    '    If IsThisApplicationInForeGround() = False Then
    '        If strApp = "LOP" And strTble = "frmBridge" Then
    '            'windowsTaskbar.SetOverlayIcon(Me.Handle, My.Resources.ModifiedIcon, "Modified")
    '            windowsTaskbar.SetOverlayIcon(Me.Handle, My.Resources.Red, "Change")
    '        End If
    '    End If
    'End Sub

    'Private Sub LOPMain_Activated(sender As System.Object, e As System.EventArgs) Handles MyBase.Activated
    '    windowsTaskbar.SetOverlayIcon(Me.Handle, Nothing, Nothing)
    'End Sub

#End Region

#Region "CreateTsb"

    '20120510 Added drop downs to the main menu.
    '20120521 Main menu is not shown as a form when coupling to a group. The sub menus are. The main menu is shown if 1 or more sub menus are visible.
    Public Function CreateMainMenuDropDownEntry(ByVal tsmParent As ToolStripMenuItem, ByVal strName As String) As ToolStripMenuItem
        Dim strTsmName = "tsm" + strName.Replace(" ", "_")
        Dim tsm As New ToolStripMenuItem

        tsm.BackColor = System.Drawing.Color.Silver
        tsm.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        tsm.ImageTransparentColor = System.Drawing.Color.Magenta

        tsm.Name = "tsm" + strName.Replace(" ", "_")
        tsm.Text = strName

        tsm.ToolTipText = strName
        AddHandler tsm.Click, AddressOf tsmMainMenu_Click
        If Not tsmParent Is Nothing Then
            tsm.AutoSize = True
            tsmParent.DropDownItems.Add(tsm)
            statics.put_v_form(strName, False, False)
        Else
            tsm.MergeAction = MergeAction.Append
            'This one means that the Maintain MenuItem acts like other buttons when there is not enough room on the FormMenuStrip.
            'The default actions is that it always remains in view shifting left when necessary.
            tsm.Overflow = ToolStripItemOverflow.AsNeeded
            tsm.AutoSize = False
            tsm.Size = New System.Drawing.Size(80, 22)
            tsm.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {})

            Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {tsm})
            ' Me.ToolStrip1.Items.Add(tsm)

            'Store the drop down menu in the forms table if this is a top-level entry.
            'Give it values 0, 1 so that it doesn't get loaded into the Form fialog when adding forms to groups.
            statics.put_v_form(strName, False, True)
        End If
        Return tsm
    End Function

    Public Function CreateTsb(ByVal strName As String, ByVal strText As String, ByVal blnAddToToolstrip As Boolean, ByVal iWidth As Integer) As ToolStripButton
        Return CreateTsb(strName, strText, blnAddToToolstrip, False, iWidth)
    End Function

    Public Function CreateTsb(ByVal strName As String, ByVal strText As String, ByVal blnAddToToolstrip As Boolean, ByVal blnAutoSize As Boolean) As ToolStripButton
        Return CreateTsb(strName, strText, blnAddToToolstrip, blnAutoSize, 100)
    End Function

    Public Function CreateTsb(ByVal strName As String, ByVal strText As String, ByVal blnAddToToolstrip As Boolean, ByVal blnAutoSize As Boolean _
     , ByVal iWidth As Integer) As ToolStripButton
        Dim tsb As ToolStripButton = New System.Windows.Forms.ToolStripButton
        tsb.Size = New System.Drawing.Size(iWidth, 25)
        tsb.AutoSize = blnAutoSize
        tsb.BackColor = System.Drawing.Color.Silver
        tsb.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        tsb.ImageTransparentColor = System.Drawing.Color.Magenta
        tsb.Name = strName
        tsb.Text = strText  'Do this in CheckFormvisibility. Also the statics are not initialised at this stage! statics.get_txt_header(strText)
        tsb.ToolTipText = strText
        If blnAddToToolstrip Then
            Me.ToolStrip1.Items.Add(tsb)
        End If
        Return tsb
    End Function

    Public Function CreateTsmForMenuStrip(ByVal strName As String, ByVal strText As String, ByVal blnAddToToolstrip As Boolean) As ToolStripMenuItem
        Dim tsm As ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        tsm.Size = New System.Drawing.Size(69, 22)
        ' tsm.AutoSize = blnAutoSize
        'tsm.BackColor = System.Drawing.Color.Silver
        tsm.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        'tsm.ImageTransparentColor = System.Drawing.Color.Magenta
        tsm.Name = strName
        tsm.Text = strText  'statics.get_txt_header(strText)
        tsm.ToolTipText = strText
        If blnAddToToolstrip Then
            Me.MenuStrip.Items.Add(tsm)
        End If
        Return tsm
    End Function

    Sub CreateDropDownMenu(ByVal tsm As ToolStripMenuItem, ByVal strName As String)
        Dim strTsmName = "tsm" + strName.Replace(" ", "_")
        tsm.Name = strTsmName
        tsm.Size = New System.Drawing.Size(128, 22)
        tsm.Text = statics.get_txt_header(strName)

        'Store the drop down menu in the forms table.
        statics.put_v_form(strName, True, False)
        'Me.tsmMasterData.BackColor = Color.AliceBlue
        tsm.ToolTipText = strName
        'tsm.AutoToolTip = False
        tsm.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {})
        Me.MenuStrip.Items.Add(tsm)
    End Sub
    Sub CreateDropDownMenuEntry(ByVal tsmParent As ToolStripMenuItem, ByVal strName As String)
        Dim tsm As New ToolStripMenuItem
        Dim strTsmName = "tsm" + strName.Replace(" ", "_")
        tsm.Name = strTsmName
        tsm.Size = New System.Drawing.Size(128, 22)
        tsm.Text = statics.get_txt_header(strName)

        tsm.ToolTipText = strName
        AddHandler tsm.Click, AddressOf tsm_Click
        If Not tsmParent Is Nothing Then
            tsmParent.DropDownItems.Add(tsm)
            statics.put_v_form(strName, False, True)
            'statics.delete_v_form(strName)
        Else
            Me.MenuStrip.Items.Add(tsm)

            'Store the drop down menu in the forms table if this is a top-level entry.
            statics.put_v_form(strName, True, True)

        End If
    End Sub

    Public Overridable Sub tsm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If blnBringToFrontIfExists(Me, sender.Text) = False Then
            If statics.blnSupportsMultiLang Then
                ShowForm(Me, statics.get_txt_header(sender.ToolTipText), sender.ToolTipText, MainDefs, Nothing)
            Else
                ShowForm(Me, sender.ToolTipText.Substring(3), sender.ToolTipText.Substring(3), MainDefs, Nothing)
            End If
        End If
    End Sub

    Protected Overridable Sub tsmMainMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ctl As ToolStripMenuItem = TryCast(sender, ToolStripMenuItem)
        If Not ctl Is Nothing Then
            If ctl.ToolTipText = SECURITY Then
                If blnBringToFrontIfExists(Me, sender.ToolTipText) = False Then
                    DisplayAForm(Me, New frmManage(sender, sender.ToolTipText, MainDefs), sender.Text, sender.ToolTipText)
                End If
            End If

            If ctl.ToolTipText = OPTIONS Then
                If blnBringToFrontIfExists(Me, sender.ToolTipText) = False Then
                    DisplayAForm(Me, New frmAppParameters(sender, sender.ToolTipText, MainDefs), sender.Text, sender.ToolTipText)
                End If
            End If

            If ctl.ToolTipText = USERLOG Then
                If blnBringToFrontIfExists(Me, sender.ToolTipText) = False Then
                    DisplayAForm(Me, New frmUsrLog(sender, sender.ToolTipText, MainDefs), sender.Text, sender.ToolTipText)
                End If
            End If

            If ctl.ToolTipText = APPLOG Then
                If blnBringToFrontIfExists(Me, sender.ToolTipText) = False Then
                    DisplayAForm(Me, New frmAppLog(sender, sender.ToolTipText, MainDefs), sender.Text, sender.ToolTipText)
                End If
            End If

            If ctl.ToolTipText = MENU_ADHOC_VIEWS Then
                If blnBringToFrontIfExists(Me, sender.ToolTipText) = False Then
                    DisplayAForm(Me, New frmAdhocViews(sender, sender.ToolTipText, MainDefs), sender.Text, sender.ToolTipText)
                End If
            End If

            If ctl.ToolTipText = MAINMENU_META_VIEWS Then
                If blnBringToFrontIfExists(Me, sender.ToolTipText) = False Then
                    DisplayAForm(Me, New frmMetaView(sender, sender.ToolTipText, MainDefs), sender.Text, sender.ToolTipText)
                End If
            End If

            If ctl.ToolTipText = MAINMENU_META_Procs Then
                If blnBringToFrontIfExists(Me, sender.ToolTipText) = False Then
                    DisplayAForm(Me, New frmMetaProc(sender, sender.ToolTipText, MainDefs), sender.Text, sender.ToolTipText)
                End If
            End If

        End If

    End Sub

    ''' <summary>
    '''     Override this sub
    ''' to create the link between the menu name and the form to be shown.
    ''' </summary>
    ''' <param name="pParent"></param>
    ''' <param name="strHeader"></param>
    ''' <param name="strForm"></param>
    ''' <param name="MainDefs"></param>
    ''' <param name="Fields"></param>
    ''' <remarks></remarks>
    Public Overridable Sub ShowForm(ByVal pParent As Form, ByVal strHeader As String, ByVal strForm As String, ByVal MainDefs As MainDefinitions, ByVal Fields As Dictionary(Of String, String))
        'idea is to create the link between the menu name and the form to be shown.
        'Dim f As GenericForm = Nothing
        'Select Case strForm.ToLower
        '    Case "b_operator"
        '        f = New form_b_operator(Nothing, strForm, MainDefs, Fields, True, False)

        'End Select
        'If Not f Is Nothing Then
        '    ut.DisplayAForm(pParent, f, strHeader, strForm.ToLower, FormWindowState.Normal)
        'End If

    End Sub

#End Region

#Region "Connection"
    Public Overridable Sub SetQualityDataBase(ByVal blnChecked As Boolean)
        'My.Settings.QualityDatabase = blnChecked
        'My.Settings.Save()
    End Sub

    Public Overridable Function GetQualityDataBase() As Boolean
        'Return My.Settings.QualityDatabase
        Return False
    End Function

    Public Overridable Function GetDataSourceLive() As String
        'Return My.Settings.DataSourceLive
        Return ""
    End Function

    Public Overridable Function GetCatalogLive() As String
        'Return My.Settings.CatalogLive
        Return ""
    End Function

    Public Overridable Function GetDataSourceQuality() As String
        'Return My.Settings.DataSourceQuality
        Return ""
    End Function

    Public Overridable Function GetCatalogQuality() As String
        'Return My.Settings.CatalogQuality
        Return ""
    End Function

    Public Overridable Function GetEnableAudio() As Boolean
        'Return My.Settings.CatalogQuality
        Return False
    End Function

    Public Overridable Sub SetEnableAudio(ByVal blnChecked As Boolean)
        'My.Settings.QualityDatabase = blnChecked
        'My.Settings.Save()
    End Sub

    Public Overridable Function GetConnectionString() As String
        Return GetConnectionString(GetQualityDataBase())
    End Function

    Public Overridable Function GetConnectionString(ByVal blnQuality As Boolean) As String
        Return My.Settings.UtilitiesConnectionString()
    End Function

    'Public MustOverride Function GetConnectionString(ByVal blnQuality As Boolean) As String

#End Region

#Region "Display_Forms"

    '20100225 Created this to improve on launching a form.
    Public Sub CloseIfExists(ByVal pParent As Form, ByVal strFormName As String)
        Dim f As Form = Nothing
        strFormName = strFormName.Trim()
        For Each f In pParent.MdiChildren()
            If f.Name = strFormName Then
                f.Close()
                Exit For
            End If
        Next
    End Sub

    Public Function blnBringToFrontIfExists(ByVal pParent As Form, ByVal strFormName As String) As Boolean
        Dim blnExists As Boolean = False
        Dim f As Form = Nothing
        Dim bFrmExists As Boolean = False
        strFormName = strFormName.Trim()
        For Each f In pParent.MdiChildren()
            If f.Name = strFormName Then
                f.BringToFront()
                blnExists = True
                Exit For
            End If
        Next
        Return blnExists
    End Function

    Public Sub DisplayAForm(ByVal pParent As Form, ByVal frm As Form, _
                ByVal strHeader As String, ByVal strFormName As String)
        DisplayAForm(pParent, frm, strHeader, strFormName, FormWindowState.Maximized)
    End Sub

    Public Sub DisplayAForm(ByVal pParent As Form, ByVal frm As Form, _
            ByVal strHeader As String, ByVal strFormName As String, ByVal frmState As FormWindowState)

        '20100105 RPB modified DisplayAForm by Trimming strHeader and strFormName
        strFormName = strFormName.Trim()
        If blnBringToFrontIfExists(pParent, strFormName) = True Then

            '20100225 
            frm.Dispose()
        Else
            strHeader = strHeader.Trim()
            Try
                ShowAForm(pParent, frm, strHeader, strFormName, frmState)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        End If
    End Sub

    Public Sub ShowAForm(ByVal pParent As Form, ByVal frm As Form, ByVal strHeader As String, _
        ByVal strFormName As String)
        ShowAForm(pParent, frm, strHeader, strFormName, FormWindowState.Maximized)
    End Sub
    Public Sub ShowAForm(ByVal pParent As Form, ByVal frm As Form, ByVal strHeader As String, _
        ByVal strFormName As String, ByVal frmState As FormWindowState)

        'Note that a form with a valid MDI parent can not be added to the Controls collection of the parent.
        Dim cur As Cursor
        cur = pParent.Cursor
        pParent.Cursor = Cursors.WaitCursor

        '20130203 Experimenting with SDI so added this.
        If Me.IsMdiContainer = True Then
            frm.MdiParent = pParent
        End If
        If strHeader.Length <> 0 Then
            frm.Text = strHeader
        End If
        frm.Name = strFormName
        frm.WindowState = frmState
        frm.ControlBox = True
        frm.Show()
        pParent.Cursor = cur

    End Sub
#End Region

#Region "Refresh"
    Public Overridable Sub RefreshServerStatusStrip(ByVal strStatus As String)
        'StatusStrip.Items("tsServerStatus").Text = strStatus
    End Sub
    Public Overridable Sub RefreshDirectoryStatusStrip(ByVal strStatus As String)
        'StatusStrip.Items("tsServerDirectory").Text = ""
    End Sub
    Public Overridable Sub RefreshStatusStrip()

        ''Try to show the initial catalog but show complete ConnectionString if it fails.
        'Try
        '    Dim strC As String = GetConnectionString()

        '    strC = strC.Substring(0, strC.IndexOf("Integrated"))
        '    Me.tsbConnection.Text = strC

        'Catch ex As Exception
        '    Me.tsbConnection.Text = GetConnectionString()
        'End Try

    End Sub
    Public Sub UpdateStatusLabel(ByVal strText As String)
        ToolStripStatusLabel.Text = strText
        ToolStripStatusLabel.Invalidate()

        'Following required to update the screen to show the new status.
        Me.Refresh()
    End Sub
#End Region

#Region "MainMenuItems"
    Private m_ChildFormNumber As Integer
    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs) Handles NewToolStripMenuItem.Click
        ' Create a new instance of the child form.
        Dim ChildForm As New System.Windows.Forms.Form
        ' Make it a child of this MDI form before showing it.
        ChildForm.MdiParent = Me
        m_ChildFormNumber += 1
        ChildForm.Text = "Window " & m_ChildFormNumber
        ChildForm.Show()
    End Sub

    Protected Overridable Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs) Handles OpenToolStripMenuItem.Click
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = OpenFileDialog.FileName
            ' TODO: Add code here to open the file.
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

        If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = SaveFileDialog.FileName
            ' TODO: Add code here to save the current contents of the form to a file.
        End If
    End Sub
    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CutToolStripMenuItem.Click
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CopyToolStripMenuItem.Click
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PasteToolStripMenuItem.Click
        'Use My.Computer.Clipboard.GetText() or My.Computer.Clipboard.GetData to retrieve information from the clipboard.
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CascadeToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticleToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileVerticalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles TileHorizontalToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ArrangeIconsToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles CloseAllToolStripMenuItem.Click
        ' Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

#End Region

#Region "To All Forms"
    Private Sub TPIMain_MdiChildActivate(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.MdiChildActivate
        Dim frmStd As frmStandard = TryCast(Me.ActiveMdiChild, frmStandard)
        If Not frmStd Is Nothing Then

            '20120709 Do not do this as resulting  FillTableAdapter() fires too early 
            'for child grids.

            '20130305 Do this but not if the form is being opened.
            If frmStd.blnAllowUpdate = True Then
                frmStd.FormIsActivated()
            End If
        End If
    End Sub

    'If UI has updated force the semaphore to read the b_semaphore table instead of waiting for the timer.
    Public Sub ForceSemaphoreRead()

        'Fails if there are no open child windows.
        Try
            CheckSemaphoreTable()
        Catch ex As Exception
        End Try
    End Sub


    ''' <summary>
    ''' A public call asks the form to refresh if it recognises the strings.
    ''' Is called from the frmSemaphore.
    ''' 20150815 added parameter lsemaphore
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Sub RefreshAll(ByVal strApp As String, ByVal strTble As String, ByVal lsemaphore As Long)

        'Fails if there are no open child windows.
        Try
            'Dim fws As FormWindowState = ActiveMdiChild.WindowState
            For Each f As Form In Me.MdiChildren()
                Try
                    Dim frmStd As frmStandard = TryCast(f, frmStandard)
                    If Not frmStd Is Nothing Then

                        '20091209 Pass the ActiveMdiChild so that form only updates if it is the active form.
                        frmStd.RefreshTheForm(Me.ActiveMdiChild, strApp, strTble, lsemaphore)
                    End If

                Catch ex As Exception
                End Try
            Next

        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' A public call which will refresh all open forms.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Sub RefreshAll()

        'Fails if there are no open child windows.
        Try
            Dim fws As FormWindowState = ActiveMdiChild.WindowState
            Dim f As Form
            For Each f In Me.MdiChildren()
                Try
                    Dim _frm = CType(f, Object)
                    _frm.RefreshTheForm()
                Catch ex As Exception
                End Try
            Next
            If Not ActiveMdiChild Is Nothing Then
                Me.ActiveMdiChild.WindowState() = fws
            End If

        Catch ex As Exception

        End Try

    End Sub
#End Region

#Region "Filters"
    Protected Overridable Sub ResetFilter()

        'Fails if there are no open child windows.
        Try
            If Not ActiveMdiChild Is Nothing Then
                Me.SuspendLayout()
                Dim fws As FormWindowState = ActiveMdiChild.WindowState
                For Each f As Form In Me.MdiChildren()
                    Try
                        Dim _frm As frmStandard = TryCast(f, frmStandard)
                        If Not _frm Is Nothing Then
                            '_frm.SuspendLayout()
                            'If ActiveMdiChild.Name <> _frm.Name Then
                            '    _frm.Visible = False
                            'End If

                            _frm.ResetFilter()
                            '_frm.Visible = True
                            '_frm.ResumeLayout()
                        End If
                    Catch ex As Exception
                    End Try
                Next
                If Not ActiveMdiChild Is Nothing Then
                    Me.ActiveMdiChild.WindowState() = fws
                End If
                Me.ResumeLayout()
            End If
        Catch ex As Exception

        End Try

    End Sub
    Protected Overridable Sub SetAFilter(ByVal sender As System.Object, _
            ByVal e As System.Windows.Forms.DataGridViewCellEventArgs, ByVal strFormName As String)

        Dim dg As DataGridView
        dg = CType(sender, DataGridView)

        'Do not call the originator of the broadcast unless form name is not defined.
        If strFormName.Length = 0 Or dg.Parent.Name <> strFormName Then
            If strFormName.Length = 0 Then strFormName = dg.Parent.Name
            Dim ctl() As Control
            ctl = Me.Controls.Find(strFormName, True)
            If ctl.Length > 0 Then
                Dim _frm = CType(ctl(0), Object)
                _frm.FilterFromOtherForm(sender, e)
            End If
        End If
    End Sub

    Public Overridable Sub BroadcastFilter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

        '20090818 RPB added this code in BroadcastFilter to support filtering between all open forms.
        Dim fws As FormWindowState = ActiveMdiChild.WindowState
        Dim dg As DataGridView
        dg = CType(sender, DataGridView)
        Dim strSender = dg.Parent.Name

        For Each f As Form In Me.MdiChildren()
            Try
                '20100308 RPB altered to TryCast.
                'Dim _frm = CType(f, Object)
                Dim _frm As frmStandard = TryCast(f, frmStandard)
                If Not _frm Is Nothing Then
                    If strSender <> _frm.Name Then
                        '_frm.SuspendLayout()
                        _frm.FilterFromOtherForm(sender, e)
                        '_frm.ResumeLayout()
                    End If
                End If
            Catch ex As Exception
            End Try
        Next
        If Not ActiveMdiChild Is Nothing Then
            Me.ActiveMdiChild.WindowState() = fws
        End If

    End Sub

    Private Sub tsbBtnResetFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbBtnResetFilter.Click
        ResetFilter()
    End Sub

#End Region

#Region "MenuButtons"

    '20131009 tsmProperties.ToolTipText instead of sender.text so the untranslated text is used.
    Private Sub tsmProperties_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmProperties.Click
        If blnBringToFrontIfExists(Me, sender.Text) = False Then
            DisplayAForm(Me, New frmProperties(Nothing, tsmProperties.ToolTipText, MainDefs), sender.Text, sender.Text)
        End If
    End Sub

    Private Sub tsmHelpText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmHelpText.Click
        If blnBringToFrontIfExists(Me, sender.Text) = False Then
            DisplayAForm(Me, New frmHelpText(Nothing, tsmHelpText.ToolTipText, MainDefs), sender.Text, sender.Text)
        End If
    End Sub

    Protected Overridable Sub SelectedFont(ByVal tsb As ToolStripItem)
        Dim SelectFont As System.Drawing.Font
        If Not tsb Is Nothing Then
            SelectFont = New System.Drawing.Font("Tahoma", 8, FontStyle.Bold Or FontStyle.Regular)
            tsb.Font = SelectFont
        End If
    End Sub


    Dim blnColumnsVisible As Boolean = False

#End Region

#Region "Thread"
    'Keep a copy of the semaphores to compare with the fresh ones.
    Dim CopySemaphoreTable As New TheDataSet.b_semaphoreDataTable
    Friend WithEvents B_semaphoreTableAdapter As TheDataSetTableAdapters.b_semaphoreTableAdapter = New TheDataSetTableAdapters.b_semaphoreTableAdapter

    Protected Overridable Sub ShowOnScreen(ByVal strApp As String, ByVal strTble As String, ByVal lsemaphore As Long)

        'Distribute the events to the other forms via the mainform.
        ' 20150815 added parameter lsemaphore
        RefreshAll(strApp, strTble, lsemaphore)
        RefreshStatusStrip()
    End Sub

    'Reads the Semaphore table and if the semaphore changes calls subs to send to all forms.
    '20110415 Changed to Protected Overridable  from Private because TPITrack needs it to be called but does not
    'call MyBase::TimerInMainThread because no autologout is needed.
    Protected Overridable Sub CheckSemaphoreTable()
        Try
            Dim semaphoreTable As TheDataSet.b_semaphoreDataTable = Me.B_semaphoreTableAdapter.GetData()

            'If this is the first time then create the copied version.
            If CopySemaphoreTable.Count <> semaphoreTable.Count Then
                If CopySemaphoreTable.Count > 0 Then
                    CopySemaphoreTable.Clear()
                End If
                For Each row As TheDataSet.b_semaphoreRow In semaphoreTable
                    CopySemaphoreTable.Addb_semaphoreRow(row.app, row.tble, row.semaphore)
                    ShowOnScreen(row.app, row.tble, 0)
                Next
            Else

                'Compare the copy and new versions of the semaphore to detect a change.
                Dim iRow As Integer = 0
                For Each row As TheDataSet.b_semaphoreRow In semaphoreTable
                    Dim r As TheDataSet.b_semaphoreRow = CopySemaphoreTable.Rows(iRow)

                    'If a difference is found generate the change 'event' and store the new value in the copy.
                    If row.semaphore <> r.semaphore Then
                        ShowOnScreen(row.app, row.tble, row.semaphore)
                        r.semaphore = row.semaphore
                    End If
                    iRow = iRow + 1
                Next
            End If
            'SQLStatus = StatusValues.OK
        Catch ex As Exception
            'SQLStatus = StatusValues.NOK
        End Try

    End Sub

    Dim blnPlayAudio As Boolean = False
    ''' <summary>
    ''' Check whether a form wants to have the audio file played. Set a flag so that the thread plays it.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Do_Audio()
        'if one child window says play the audio do it and leave.
        If strAudioFile.Length > 0 Then
            'Fails if there are no open child windows.
            Try
                blnPlayAudio = False
                For Each f As Form In Me.MdiChildren()
                    Try
                        Dim frmStd As frmStandard = TryCast(f, frmStandard)
                        If Not frmStd Is Nothing Then
                            If frmStd.blnPlayAudio() = True Then
                                blnPlayAudio = True
                                'My.Computer.Audio.Play(strAudioFile)
                                Exit For
                            End If
                        End If
                    Catch ex As Exception
                    End Try
                Next
            Catch ex As Exception
            End Try
        End If
    End Sub


    Dim WorkerThread As Thread
    Dim blnWorkerThreadRun As Boolean = True

    'Semaphore and TPITrack use these so careful when changing.
    Const iSLEEPSECONDS = 5    '100
    Dim StartTime As DateTime
    Dim SpanToGo As TimeSpan
    Dim EndTime As DateTime = DateTime.Now
    Dim iLongSleep As Integer = iSLEEPSECONDS * 1000
    Dim iShortSleep As Integer = 1000

    'The thread actually sleeps this long so that it can end quickly.
    Const BackGroundSleep = 250
    Dim iAutoLogout As Integer = -1
    Dim iAutoLogoutWarning As Integer = -1
    Public Delegate Sub WorkerThreadCaller()
    Public Delegate Sub UpdateUICB(ByVal strMessage As String, ByVal blnEndOfProcess As Boolean)
    Private Sub StartTimer()
        StartTime = DateTime.Now
        Dim strAutoLogout As String = ""

        'It actually fails here if the b_app_parameter is not in the database.
        Try
            strAutoLogout = statics.GetParameter("auto_logout")
            If strAutoLogout = "" Then
                iAutoLogout = -1
            Else
                iAutoLogout = System.Convert.ToInt32(strAutoLogout)
                If iAutoLogout > 0 Then
                    EndTime = StartTime.AddMinutes(iAutoLogout)
                    SpanToGo = New TimeSpan(EndTime.Ticks - DateTime.Now.Ticks)
                    Me.tslTimeToGo.Text = ": " + System.Convert.ToInt32((SpanToGo.TotalMinutes)).ToString() & " minutes"
                    ToolStrip1.Refresh()
                End If
            End If

        Catch ex As Exception
            iAutoLogout = -1
        Finally
        End Try
        If Not WorkerThread Is Nothing Then
        Else
            WorkerThread = New Thread(AddressOf RunsOnWorkerThread)
            Me.B_semaphoreTableAdapter.ClearBeforeFill = True
            Me.B_semaphoreTableAdapter.Connection.ConnectionString = MainDefs.MainForm.GetConnectionString()
            WorkerThread.Start()
        End If
    End Sub

    'Stop the thread and wait to give it a chance to end correctly.
    Private Sub StopThread()
        Dim iCount As Integer = 10
        If Not WorkerThread Is Nothing Then
            blnWorkerThreadRun = False
            While WorkerThread.IsAlive And iCount > 0
                Thread.Sleep(100)
                iCount = iCount - 1
            End While
            WorkerThread = Nothing
        End If
    End Sub

    'called from the thread when the timer fires.
    Private Sub TimerOutSideThread(ByVal strTypeOfTick As String, ByVal blnEndOfProcess As Boolean)
        Dim returnValue As IAsyncResult
        If strTypeOfTick <> strSuperShortTimerTick Then
            Try
                returnValue = Me.BeginInvoke(New UpdateUICB(AddressOf TimerInMainThread), strTypeOfTick, blnEndOfProcess)
                Thread.Sleep(0)

                '20110415 Then wait until the asynch call comes back.
                'Not doing this can cause a stack overflow if the TimerInMainThread repeatedly takes a long time to run.
                If Not Me Is Nothing Then
                    Try
                        If returnValue.IsCompleted = False Then
                            Me.EndInvoke(returnValue)
                        End If
                    Catch ex As Exception
                    End Try
                Else
                    'MsgBox("the application already nothing")
                End If
            Catch ex As Exception

            End Try
        End If

        'call the TimerOutSideThread in each open form.
        For Each f As Form In Me.MdiChildren()
            Try
                Dim frmStd As frmStandard = TryCast(f, frmStandard)
                If Not frmStd Is Nothing Then
                    frmStd.TimerOutSideThread()
                End If
            Catch ex As Exception
            End Try
        Next

    End Sub

    'Timer outside the timer thread (and therefore in main thread).
    Protected Overridable Sub TimerInMainThread(ByVal strTypeOfTick As String, ByVal blnEndOfProcess As Boolean)

        'Called from worker thread via begininvoke.
        If blnEndOfProcess = False Then
            If strTypeOfTick = iShortTimerTick Then
                Do_Audio()
            Else
                If strTypeOfTick = iTimerTick Then
                    CheckSemaphoreTable()

                    UpdateRamState()
                    ToolStrip1.Refresh()

                    'call the TimerInMainThread in each open form.
                    For Each f As Form In Me.MdiChildren()
                        Try
                            Dim frmStd As frmStandard = TryCast(f, frmStandard)
                            If Not frmStd Is Nothing Then
                                frmStd.TimerInMainThread()
                            End If
                        Catch ex As Exception
                        End Try
                    Next
                    'AUTO LOGOUT processing
                    If iAutoLogout <> -1 Then

                        'Get the parameter again to enable online modification.
                        Dim strAutoLogout As String = ""
                        strAutoLogout = statics.GetParameter("auto_logout")
                        Dim strAutoLogoutWarning As String = statics.GetParameter("auto_logout_warning")
                        Try



                            'time in minutes.
                            If iAutoLogout <> System.Convert.ToInt32(strAutoLogout) Then

                                'if the logout time changes then restart the timer.
                                iAutoLogout = System.Convert.ToInt32(strAutoLogout)
                                StartTime = DateTime.Now
                            End If

                            If iAutoLogout > 0 Then

                                'Warning about stopping application. Convert minutes to seconds
                                Try
                                    iAutoLogoutWarning = 60 * System.Convert.ToInt32(strAutoLogoutWarning)
                                Catch ex As Exception
                                    iAutoLogoutWarning = 60 * 3
                                End Try
                                EndTime = StartTime.AddMinutes(iAutoLogout)

                                '20101230 Make warning come up only after 50 seconds because otherwise it is not possible to
                                'adjust the iAutoLogout.
                                If iAutoLogoutWarning >= iAutoLogout * 60 Then
                                    iAutoLogoutWarning = iAutoLogout * 60 - 50
                                End If

                                'Time to go
                                SpanToGo = New TimeSpan(EndTime.Ticks - DateTime.Now.Ticks)
                                Dim secs As Integer = SpanToGo.TotalSeconds
                                If secs < (iAutoLogoutWarning + 1) Then
                                    CloseAllForms()

                                    'Switch on Restart controls.
                                    Me.tbAppWillEnd.Visible = True
                                    Me.tbEndSeconds.Visible = True
                                    Me.tbWillEndSeconds.Visible = True
                                    Me.tbEndSeconds.Text = secs.ToString()
                                    btnRestart.Visible = True
                                End If
                                Me.tslTimeToGo.Text = ": " + System.Convert.ToInt32((SpanToGo.TotalMinutes)).ToString() & " minutes"
                                ToolStrip1.Refresh()
                                If secs <= 0 Then
                                    statics.Closing()
                                    Application.Exit()
                                End If
                            Else
                                Me.tslTimeToGo.Text = ""
                                ToolStrip1.Refresh()
                            End If
                            SQLStatus = statics.StatusValues.OK
                        Catch ex As Exception
                            iAutoLogout = -1
                            SQLStatus = statics.StatusValues.NOK
                        End Try
                    End If
                End If
            End If
        Else
            SQLStatus = statics.StatusValues.NOK
        End If
    End Sub

    'user wants to restart the timer to prevent automatic ending of the executable.
    Private Sub btnRestart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRestart.Click
        'Switch off Restart controls.
        btnRestart.Visible = False
        Me.tbAppWillEnd.Visible = False
        Me.tbEndSeconds.Visible = False
        Me.tbWillEndSeconds.Visible = False
        StartTime = DateTime.Now
        TimerInMainThread(iTimerTick, False)

    End Sub

    'this is the thread.
    Private Sub RunsOnWorkerThread()
        Dim iLongCount As Integer = 0
        Dim iShortCount As Integer = 0
        Try
            Thread.Sleep(BackGroundSleep)   'iSleep)
            While (blnWorkerThreadRun)

                iLongCount = iLongCount + BackGroundSleep
                iShortCount = iShortCount + BackGroundSleep
                If iShortCount > iShortSleep Then

                    'Fires every second.
                    TimerOutSideThread(iShortTimerTick, False)
                    If blnPlayAudio = True Then
                        My.Computer.Audio.Play(strAudioFile)
                    End If
                    iShortCount = 0
                Else
                    TimerOutSideThread(strSuperShortTimerTick, False)
                End If

                'Fires every 5 seconds.
                If iLongCount > iLongSleep Then
                    TimerOutSideThread(iTimerTick, False)
                    iLongCount = 0
                End If
                Thread.Sleep(BackGroundSleep)
                ' Debug.Print("in the thread")
            End While
        Catch ex As Exception
            TimerOutSideThread(ex.Message, False)
        Finally
            'TimerOutSideThread("Ended", True)
        End Try
    End Sub
#End Region

#Region "MaxScreen"

    '20120510 Added to switch off controls to maximise the form.
    Public Sub SwitchControls(ByVal blnOn As Boolean)
        Me.MenuStrip.Visible = blnOn
        Me.ToolStrip1.Visible = blnOn
        Me.StatusStrip.Visible = blnOn
    End Sub

#End Region

    Private Sub MenuStrip_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MenuStrip.KeyDown
        If e.KeyCode = Keys.F10 Then
            ResetFilter()
        End If
    End Sub

    'Private Sub MainForm_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
    '    Debug.Print(e.KeyChar.ToString)
    'End Sub

    'Private Sub MainForm_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
    '    Debug.Print(e.KeyCode.ToString)
    'End Sub
    'Protected Overrides Function ProcessDialogKey(ByVal keyData As Keys) As Boolean
    '    If (keyData = 262162) Then
    '        Debug.Print("The 262162 alt" & keyData.ToString)
    '        Return True
    '    End If
    '    Debug.Print(keyData.ToString)
    '    Return MyBase.ProcessDialogKey(keyData)
    'End Function

    'This is the way to disable the close button of the form.
    'Needs to be called from ReSize see below.
    'Private Const SC_CLOSE As Integer = &HF060
    'Private Const SC_RESTORE As Integer = &HF120
    'Private Const MF_GRAYED As Integer = &H1
    'Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
    'Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
    'Private Sub frmStandard_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
    '    EnableMenuItem(GetSystemMenu(Me.Handle, False), SC_RESTORE, MF_GRAYED)
    '    Debug.Print("mainform Activated")
    'End Sub

    'Const MF_BYCOMMAND As UInteger = &H0
    'Const MF_ENABLED As UInteger = &H0
    'Const WM_SHOWWINDOW As Integer = &H18
    'Const WM_CLOSE As Integer = &H10

End Class
