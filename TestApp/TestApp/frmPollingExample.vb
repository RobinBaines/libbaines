'------------------------------------------------
'Name: Module frmPollingExample.vb
'Function: Show use of the TimerInMainThread() and RefreshTheForm() calls.
'Copyright Baines 2012. All rights reserved.
'Notes: Is derived from stdDialog.  
'The TimerInMainThread() is called every 5 seconds.
'A semaphore can be used to trigger an event in a form. In this example the semaphore writes a message in the right hand listbox. 

'Start by inserting the TESTAPP HEARTBEAT record in b_semaphore from MSSQLS:
'INSERT INTO [dbo].[b_semaphore]([app] ,[tble], [semaphore])SELECT 'TestApp', 'Heartbeat' , 1 WHERE NOT EXISTS (SELECT 1 FROM [b_semaphore] WHERE [app] = 'TestApp' AND tble = 'Heartbeat' 

'Then increment it to trigger the message in the right hand listbox.
'UPDATE [b_semaphore] SET semaphore = semaphore + 1 WHERE [app] = 'TestApp' AND tble = 'Heartbeat' 

'Modifications: 
'------------------------------------------------
Imports Utilities
Imports System
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Threading
Imports ExcelInterface.XMLExcelInterface
Imports System.ComponentModel
Public Class frmPollingExample
    Dim blnRun As Boolean = False
#Region "New"
    Public Sub New()
        MyBase.New()
        InitializeComponent()
    End Sub

    Public Sub New(ByVal tsb As ToolStripItem _
               , ByVal strSecurityName As String, ByVal _MainDefs As MainDefinitions)
        MyBase.New(tsb, strSecurityName, _MainDefs)
        InitializeComponent()
        blnRun = True
    End Sub
#End Region
#Region "Timer"
    Public Overrides Sub TimerInMainThread()
        MyBase.TimerInMainThread()
        If blnRun = False Then Exit Sub
        statics.UpdateAppLog("PollingExample Called", TestStatics.ErrorLevels.LowPriority)
        ListBox1.Items.Add(Date.Now.ToString + "PollingExample Called added from TimerInMainThread")
    End Sub
#End Region

    'This is called when the semaphore field in a record in b_semaphore changes.
    Public Overrides Sub RefreshTheForm(ByVal activeForm As Form, ByVal strApp As String, ByVal strTble As String, ByVal lsemaphore As Long)
        If strApp.ToUpper = "TESTAPP" And strTble.ToUpper = "HEARTBEAT" Then
            blnRefreshIsNeeded = True
            If Not activeForm Is Nothing Then
                If activeForm.Name = Me.Name Then
                    ListBox2.Items.Add(Date.Now.ToString + " b_semaphore event received " + strApp + "." + strTble)
                End If
            End If
        End If
    End Sub

End Class