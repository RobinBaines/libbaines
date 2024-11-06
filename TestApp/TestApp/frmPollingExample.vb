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
Imports ClosedXML.Excel
Imports System.Linq
Public Class frmPollingExample
    Dim blnRun As Boolean = False
    Friend WithEvents tsbImportReceiptOrder As ToolStripButton
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
        Me.tsbImportReceiptOrder = Me.CreateTsb("tsbImportReceiptOrder", "Import Receipt Order", True, True, 80)
        Me.SwitchOffPrintDetail()
        Me.SwitchOffPrint()
        Me.SwitchOffUpdate()
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

    Protected Sub tsbImportReceiptOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbImportReceiptOrder.Click

        Dim fileDlg As New OpenFileDialog
        fileDlg.Title = fileDlg.Title
        fileDlg.Filter = "Receipt Order File (*.xls*)|*.xls*"
        fileDlg.DefaultExt = "xlsx"
        fileDlg.InitialDirectory = statics.GetParameter("IMPORT")
        If fileDlg.InitialDirectory.Length = 0 Then
            fileDlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            fileDlg.InitialDirectory = "c:\Projects\spits\GIS\RECEIPT_ORDERs\"
        End If
        If fileDlg.ShowDialog() = DialogResult.OK Then
            Try
                Dim ReadExcel As New ExcelInterface.ReadExcel
                Dim A As Collection
                A = ReadExcel.Open(fileDlg.FileName)   '"\\RPB5\Projects\spits\GIS\receipt.xlsx"
                If A.Count > 1 Then
                    '                        V_drugid__receiptTableAdapter.p_create_a_dotinfo_ro(strSupplier, iRo_id)
                    Dim strArray(8)
                    Dim iRowCount As Integer = 1
                    For Each strArray In A
                        If strArray.Count() > 8 Then
                            Dim strZinr As String = strArray(0)
                            If strZinr Is Nothing Then
                                Exit For
                            End If
                            If strZinr.Length > 7 Then
                                strZinr = strZinr.Substring(0, 8)
                                Dim strPackCount As String = strArray(7)
                                Dim iPackCount As Integer
                                Try
                                    iPackCount = System.Convert.ToInt32(strPackCount)
                                    'If strZinr.Contains("/") Then
                                    '    strZinr = strZinr.Substring(0, strZinr.IndexOf("/"))
                                    'End If

                                    Dim iRowAddedCount As Integer = 0
                                    '     V_drugid__receiptTableAdapter.p_create_a_ro_line(iRo_id, strZinr, iPackCount, iRowAddedCount)
                                Catch ex As Exception

                                    MsgBox(iRowCount.ToString & statics.get_txt_header(" row is a problem. Check format.",
          "User information raised when creating receipt order lines in ...", "User information"))
                                End Try
                            End If
                        End If
                        iRowCount += 1
                    Next
                End If

            Catch ex As Exception
                MsgBox(statics.get_txt_header("Not created from Excel file. " + ex.Message, _
                          "User information raised when creating lines in ...", "User information"))
                End Try
            Else
                ' the dialog was cancelled  
            End If

    End Sub
End Class