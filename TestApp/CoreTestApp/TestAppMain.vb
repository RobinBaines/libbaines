'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created Jan 2008.
'This Main form is derived from Utilities.MainForm.

'Notes: The Namespace of this app is TestApp.
'If there is a class also called TestApp then the designer will add TestApp to TestApp.TheDataSet and compiler thinks this a references to 
'the TestApp class while it is the NameSpace.
'So do not name a class TestApp!!
'Modifications:
'------------------------------------------------
Imports System
Imports System.Security.Cryptography
Imports System.Configuration

Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Sql
Imports Utilities
Public Class TestAppMain

    'The MenuStrip is the top row of menus. These tend to be used for Master data.
    'The second row of menu buttons is the ToolStrip1 and is used as the main menu.
    'This is a button which opens a form and is put in the top level of the ToolStrip1
    Protected WithEvents tsbfrmTest As System.Windows.Forms.ToolStripButton
    Const MENU_TESTFORM As String = "TestForm"

    Protected WithEvents tsbfrmFlow As System.Windows.Forms.ToolStripButton
    Const MENU_FLOWFORM As String = "FlowForm"

    'This is a drop down called 'Some Tests' with 2 items which open forms.
    Const MENU_TEST As String = "Some Tests"
    Const MENU_TEST_TESTFORM As String = "2nd TestForm"
    Const MENU_TEST_POLLING As String = "Polling"
    Protected WithEvents tsbSomeTests As System.Windows.Forms.ToolStripMenuItem

    'Public SQLParser As clParseSQL

#Region "New"
    Public Sub New()

        'set the multilanguage flag 
        MyBase.New(True)
        InitializeComponent()

    End Sub
#End Region
#Region "Load"
    Public Overrides Sub Init()

        ' Me.IsMdiContainer = False
        'Create the menus before calling init() so that their visibility can be adjusted.
        'CreateDropDownMenus()
        MyBase.Init()

        'switch off some buttons in the main form.
        SwitchOffFile()
        SwitchOffEdit()
        SwitchOffTools()
        SwitchOffView()
        ''SwitchOffWindows()
        SwitchOffHelp()
        ' SQLParser = New clParseSQL()
    End Sub

    'Possible to change the app details which are shown bottom left. The default is shown here and so commented out.
    'Protected Overrides Function strGetAppDetails() As String
    '    Return Application.ProductVersion & " of " & Application.ProductName & " by " & Application.CompanyName
    'End Function

    'The ToolStrip1 with the main menu items raises this event when a drop down button is clicked.
    'Remember the ToolTipText has the original name of the button while the text of the button may have been modified for the users.
    Protected Overrides Sub tsmMainMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.tsmMainMenu_Click(sender, e)
        Dim ctl As ToolStripMenuItem = TryCast(sender, ToolStripMenuItem)
        If Not ctl Is Nothing Then
            If blnBringToFrontIfExists(Me, sender.ToolTipText) = False Then
                Select Case ctl.ToolTipText
                    Case MENU_TEST_TESTFORM
                        ' ShowAForm(Me, New frmTest(sender, sender.ToolTipText, MainDefs), sender.Text, sender.ToolTipText)
                    Case MENU_TEST_POLLING
                        ' ShowAForm(Me, New frmPollingExample(sender, sender.ToolTipText, MainDefs), sender.Text, sender.ToolTipText)
                    Case MENU_ADHOC_VIEWS
                        ShowAForm(Me, New frmAdhocViews(sender, sender.ToolTipText, MainDefs), sender.Text, sender.ToolTipText)
                End Select
            End If
        End If
    End Sub

    'Here is ToolStrip1 click event for the top level button.
    'frmTest is being open from here and from the drop down. 
    Private Sub tsbfrmTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbfrmTest.Click
        If blnBringToFrontIfExists(Me, sender.ToolTipText) = False Then
            ShowAForm(Me, New SQLBrowser(sender, sender.ToolTipText, MainDefs), sender.Text, sender.ToolTipText)
            '  ShowAForm(Me, New HTMLView(), sender.Text, sender.ToolTipText)
        End If
    End Sub

    Private Sub tsbfrmFlow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbfrmFlow.Click
        If blnBringToFrontIfExists(Me, sender.ToolTipText) = False Then
            ShowAForm(Me, New frmFlow(sender, sender.ToolTipText, MainDefs), sender.Text, sender.ToolTipText)
        End If
    End Sub

    'These are some MenuStrip items which do not open anything but are useful for testing.
    Friend WithEvents tsmMasterData As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsmb_operators As System.Windows.Forms.ToolStripMenuItem

    'CreateDropDownMenus is called during initialisation and places the ToolStrip1 items in front of the Security drop down.
    'The user (windows login) is checked before this is called and if they are blocked or not in the list of users 
    'the items should not be added.
    'If they are added but the user does not have rights to use then they will not be shown.
    Protected Overrides Sub CreateDropDownMenus()

        'If blocked then user is not allowed to see anything so dont show the default form.
        If statics.Blocked = False And statics.IsAuthorized = True Then

            'Coupling to a group is different depending on whether the top menu or the main menu.

            'MenuStrip items (the top most menu).
            'This creates a top level entry which opens a form.
            Me.tsmb_operators = New System.Windows.Forms.ToolStripMenuItem
            CreateDropDownMenu(tsmb_operators, "Operators") 'add to group to activate

            'This creates a drop down.
            Me.tsmMasterData = New System.Windows.Forms.ToolStripMenuItem
            CreateDropDownMenu(tsmMasterData, "Parent") 'add to group to activate all children
            CreateDropDownMenuEntry(tsmMasterData, "Child") 'Shows automatically if parent is activated.

            'ToolStrip items. ToolStrip1 is the main menu.
            'This creates a top level entry which opens a form.
            tsbfrmTest = Me.CreateTsb("tsbfrmTest", MENU_TESTFORM, True, False) 'add to group to activate.

            tsbfrmFlow = Me.CreateTsb("tsbfrmFlow", MENU_FLOWFORM, True, False)

            'This creates a drop down.
            tsbSomeTests = CreateMainMenuDropDownEntry(Nothing, MENU_TEST) 'Shows automatically if at least one child is activated.
            CreateMainMenuDropDownEntry(tsbSomeTests, MENU_TEST_TESTFORM) 'add to group to activate.
            CreateMainMenuDropDownEntry(tsbSomeTests, MENU_TEST_POLLING) 'add to group to activate.

        End If
    End Sub
#End Region
#Region "ConnectionStrings"

    'This is not used but retained as background.
    Shared Sub ToggleConfigEncryption(ByVal exeConfigName As String)
        ' Takes the executable file name without the
        ' .config extension.
        Try
            ' Open the configuration file and retrieve 
            ' the connectionStrings section.
            Dim config As Configuration = ConfigurationManager. _
                OpenExeConfiguration(exeConfigName)

            Dim section As ConnectionStringsSection = DirectCast( _
                config.GetSection("connectionStrings"),  _
                ConnectionStringsSection)

            If section.SectionInformation.IsProtected Then
                ' Remove encryption.
                section.SectionInformation.UnprotectSection()
            Else
                ' Encrypt the section.
                'DataProtectionConfigurationProvider uses the Windows DPAPI to perform encryption. 
                'The RSAProtectedConfigurationProvider uses the public-key algorithm available in the .NET 
                'Framework() 's RSACryptoServiceProvider class to perform encryption.
                'For example, if you use the Windows DataProtectionConfigurationProvider, 
                'the decryption key is auto-generated and saved in the Windows Local Security Authority (LSA). 

                'User-level RSA key containers are stored with the Windows user profile for a particular user 
                'and can be used to encrypt and decrypt information for applications that run under that specific user identity. 
                'User-level RSA key containers can be useful if you want to ensure that the RSA key information is removed when 
                'the Windows user profile is removed. However, because you must be logged in with the specific user account 
                'that will make use of the user-level RSA key container in order to encrypt or decrypt protected configuration 
                'sections, they are inconvenient to use.

                'Machine-level RSA key containers are available to all users that can log in to a computer, 
                'by default, and are the most useful as you can use them to encrypt or decrypt protected configuration sections 
                'while logged in with an administrator account. A machine-level RSA key container can be used to 
                'protect information for a single application, all the applications on a server, 
                'or a group of applications on a server that run under the same user identity. 
                'Although machine-level RSA key containers are available to all users, they can be secured with NTFS Access 
                'Control Lists (ACLs) so that only required users can access them.

                'Am using machine level because OpenExeConfiguration(exeConfigName) specifies the config file in exe directory.
                'There is an overload to specify user level.
                'Machine level Key is created in C:\Documents and Settings\All Users\Application Data\Microsoft\Crypto\RSA\MachineKeys
                'aspnet_regiis -px NetFrameworkConfigurationKey key.xml to export.
                '<RSAKeyValue><Modulus>27GDpg2vkomxIb1kH2JlNgUwHFQ55VYXLIAm5Wi5TX0pWrA2Yk7BK8UolNZB/zrywVDNNcIg84x7j2PYkI1IELYbS3PUSI4HtqtDg9kGkDirb/79fH3GQJ9WmKD2NVI2KAz6r2I7kC8ttXajsv/hojqVq4mC0PuPHbgn7k2Hd38=
                '</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>
                section.SectionInformation.ProtectSection( _
                "RsaProtectedConfigurationProvider")
                '  "DataProtectionConfigurationProvider")
            End If

            ' Save the current configuration.
            config.Save()

            Console.WriteLine("Protected={0}", _
            section.SectionInformation.IsProtected)

        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
    Public Overrides Sub SetQualityDataBase(ByVal blnChecked As Boolean)
        My.Settings.TestDataBase = blnChecked
        My.Settings.Save()
    End Sub
    Public Overrides Function GetQualityDataBase() As Boolean
        Return My.Settings.TestDataBase
    End Function
    Public Overrides Function GetDataSourceLive() As String
        Return My.Settings.DataSourceLive
    End Function
    Public Overrides Function GetCatalogLive() As String
        Return My.Settings.CatalogLive
    End Function
    Public Overrides Function GetDataSourceQuality() As String
        Return My.Settings.DataSourceTest
    End Function
    Public Overrides Function GetCatalogQuality() As String
        Return My.Settings.CatalogTest
    End Function
    Public Overrides Function GetEnableAudio() As Boolean
        Return My.Settings.EnableAudio
    End Function
    Public Overrides Sub SetEnableAudio(ByVal blnChecked As Boolean)
        My.Settings.EnableAudio = blnChecked
        My.Settings.Save()
    End Sub

    'When developing the connection string will be one defined in the Settings.
    'In this case 'Data Source=RPB4;Initial Catalog=Utilities;Integrated Security=True'
    'To avoid having to alter config files when going live when connections will be different the following call must be made
    'for all table adapters. It constructs the connection string depending on whether the development computer is being used
    'and on whether the test database (aka Quality database) is being used.
    'A common source of error is not to set the connection string of an adapter using this call when making a connection; 
    'that works fine on the development computer but fails live!!
    'See TimerInMainThread below for a case where this is done correctly.
    Public Overrides Function GetConnectionString(ByVal blnQuality As Boolean) As String
        'Dim strRet As String = ""

        'Dim strDataSource As String
        'Dim strCatalog As String = ""
        ''strRet = My.Settings.SQLUser
        'strRet = My.Settings.ConnectionString
        'If My.Computer.Name = My.Settings.DataSourceDevelopment Then
        '    If blnQuality = True Then
        '        strDataSource = My.Settings.DataSourceDevelopmentTest
        '    Else
        '        strDataSource = My.Settings.DataSourceDevelopment
        '    End If
        '    strCatalog = My.Settings.CatalogDevelopment
        'Else
        '    If blnQuality = True Then
        '        strDataSource = My.Settings.DataSourceTest
        '        strCatalog = My.Settings.CatalogTest
        '    Else
        '        strDataSource = My.Settings.DataSourceLive
        '        strCatalog = My.Settings.CatalogLive
        '    End If
        'End If
        'strRet = strRet.Replace("Data Source=" + My.Settings.DataSourceDevelopment, "Data Source=" & strDataSource)
        'strRet = strRet.Replace("Initial Catalog=" + My.Settings.CatalogDevelopment, "Initial Catalog=" & strCatalog)
        'Return strRet

        'check for quality as this can change on the fly.
        If ConnectionString.ConnectionString.Length = 0 Or blnQuality <> ConnectionString.Quality Then
            'Windows authentication.
            ConnectionString.Init(False, _
                       My.Settings.DataSourceDevelopment, _
                        My.Settings.DataSourceDevelopmentTest, _
                        My.Settings.CatalogLive, _
                        My.Settings.DataSourceLive, My.Settings.CatalogLive, _
                        My.Settings.DataSourceLive, My.Settings.CatalogLive)
            'End If
        End If
        Return ConnectionString.ConnectionString
    End Function
#End Region

#Region "FromThread"
    'The MainForm parent has a 5 second timer. 
    'This may be used to call the frmStandard.TimerInMainThread() whihc can be overriden.
    'frmPollingExample() illustrates this.

    'This example also illustrates the use of the b_semaphore table for passing messages.
    'Is very simple and with SQL 2008/12 there may be betters ways of doing. 
    'MainForm polls b_semaphore and passes changes to all forms. 
    'So in this case TestApp updates b_semaphore and another application will see this and could use as a heartbeat check that TestApp is running. 
    Friend WithEvents b_semaphoreTableAdapter As TheDataSetTableAdapters.b_semaphoreTableAdapter = Nothing

    ' The name TimerInMainThread means that this call is made in the main thread and not in the timer thread.
    'There is also a timer thread call but this may not be used for the UI.

    'TimerInMainThread fires every Const iSLEEPSECONDS = 5 (see Utilities.MainForm) 
    'Use this to slow it down a bit.
    Dim iCount As Integer = 0
    Const COUNT_TO = 3

    Protected Overrides Sub TimerInMainThread(ByVal strTypeOfTick As String, ByVal blnEndOfProcess As Boolean)
        MyBase.TimerInMainThread(strTypeOfTick, blnEndOfProcess)

        If b_semaphoreTableAdapter Is Nothing Then
            b_semaphoreTableAdapter = New TheDataSetTableAdapters.b_semaphoreTableAdapter
            b_semaphoreTableAdapter.Connection.ConnectionString = GetConnectionString()
        End If

        'Do not call MyBase. This disables audio, b_semaphore check and auto_logout 
        'and because the auto logout parameter applies to TPINet and not to TPITrack.
        'MyBase.TimerInMainThread(strTypeOfTick, blnEndOfProcess)
        If strTypeOfTick = iTimerTick Then
            If iCount >= COUNT_TO Then
                For Each f As Form In Me.MdiChildren()
                    Try
                        Dim frmStd As frmStandard = TryCast(f, frmStandard)
                        If Not frmStd Is Nothing Then
                            frmStd.TimerInMainThread()
                        End If
                    Catch ex As Exception
                    End Try
                Next
                'Following assumes "TestApp", "heartbeat" is in b_semaphore: INSERT INTO [Utilities].[dbo].[b_semaphore]([app],[tble],semaphore)VALUES('TestApp','heartbeat',0)
                'UpdateQuery increments semaphore counter.

                'b_semaphoreTableAdapter.UpdateQuery("TESTAPP", "HEARTBEAT")
            End If
            iCount = 0
        End If
        iCount = iCount + 1
    End Sub
#End Region
End Class
