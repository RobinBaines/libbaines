'------------------------------------------------
'Name: Module clPrintLabels.vb
'Function: 
'Copyright Baines 2012. All rights reserved.
'Notes:
'pText.AppendLine("^XA")
'pText.AppendLine("^FO150,150") 'X, Y where Y is distance from start of label
'pText.AppendLine("^BXR,10,200") 'BX is Data matrix, N is normal, R rotate 90 degrees, height =10, quality level = 200 = Reed Solomon
''pText.AppendLine("^FWR")    'rotates 90 degrees
'pText.AppendLine("^FD" & dgParent.CurrentRow.Cells("pwd").Value)
'pText.AppendLine("^FS")
''ADN is font with size 36,20, N means do not rotate R= 90 degrees
''FD is field data
'pText.AppendLine("^FO100,400^ADR,36,20^FD" & dgParent.CurrentRow.Cells("operator").Value & "^FS")
'pText.AppendLine("^XZ")

'this works too (all label in 1 line)
'pText.AppendLine("^XA^FO150,150^BXR,10,200^FD" & dgParent.CurrentRow.Cells("pwd").Value & "^FS^FO100,400^ADR,36,20^FD" & dgParent.CurrentRow.Cells("operator").Value & "^FS^XZ")

'RawPrinter.PrintZPL("169.254.168.180", pText.ToString)

'RawPrinterLocal.PrintRaw("GK420d", pText.ToString)

'Modifications:

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Net.Sockets
Imports System.Text

Public Class PrintLabels
    Dim taLabelType As PrintLabelDataSetTableAdapters.v_labeltypeTableAdapter
    Dim taLabelType_definition As PrintLabelDataSetTableAdapters.b_labeltype_definitionTableAdapter
    Dim taLabelTypeTable As PrintLabelDataSet.v_labeltypeDataTable
    Dim taLabelType_definitionTable As PrintLabelDataSet.b_labeltype_definitionDataTable
    Dim LabelName As String = ""
    Private blnCanPrint As Boolean = False
    Public ReadOnly Property CanPrint() As Boolean
        Get
            Return blnCanPrint
        End Get
    End Property

    Public Sub New(_LabelName As String)
        LabelName = _LabelName
        taLabelType = New PrintLabelDataSetTableAdapters.v_labeltypeTableAdapter
        taLabelType.Connection.ConnectionString = ConnectionString.ConnectionString
        taLabelType_definition = New PrintLabelDataSetTableAdapters.b_labeltype_definitionTableAdapter
        taLabelType_definition.Connection.ConnectionString = ConnectionString.ConnectionString
        taLabelTypeTable = New PrintLabelDataSet.v_labeltypeDataTable
        taLabelType_definitionTable = New PrintLabelDataSet.b_labeltype_definitionDataTable
        FillTableAdapter(False)
    End Sub

    Private Sub FillTableAdapter(blnShowMessage As Boolean)
        'get the printer adres.
        taLabelTypeTable = taLabelType.GetDataBy(LabelName)

        'get the label file.
        taLabelType_definitionTable = taLabelType_definition.GetDataBy(LabelName)
        blnCanPrint = CanWePrint(False)
    End Sub

    Public Function CanWePrint(blnShowMessage As Boolean) As Boolean
        If taLabelTypeTable.Count <> 1 Then
            If blnShowMessage Then
                MsgBox(statics.get_txt_header("No printer defined for this operation. " & LabelName, _
                                      "Configuration error raised when printing a label from clPrintLabels..", "configuration error"))
            End If
            Return False
        End If
        If taLabelType_definitionTable.Count = 0 Then
            If blnShowMessage Then
                MsgBox(statics.get_txt_header("No label definition for this operation. " & LabelName, _
                                 "Configuration error raised when printing a label from clPrintLabels.", "configuration error"))
            End If
            Return False
        End If
        Return True
    End Function

    'return the label definition so that caller can replace any parameters.
    Public Function GetLabelDefinition() As String
        If blnCanPrint Then
            Dim pText As New StringBuilder
            For Each row As PrintLabelDataSet.b_labeltype_definitionRow In taLabelType_definitionTable.Rows
                pText.AppendLine(row.line)
            Next
            Return pText.ToString
        End If
        Return ""
    End Function

    'after caller has replaced the parameters print it.
    Public Function PrintLabel(strLabel As String) As Boolean
        Dim blnRet As Boolean = False
        If blnCanPrint Then
            Dim rp As PrintLabelDataSet.v_labeltypeRow = taLabelTypeTable.Rows(0)
            Try
                If rp.ip = True Then
                    blnRet = PrinterIP.PrintZPL(rp.adres, strLabel.ToString)
                Else
                    blnRet = PrinterLocal.PrintRaw(rp.adres, strLabel.ToString)
                End If
            Catch ex As Exception
            End Try
        End If
        Return blnRet
    End Function
End Class


''' <summary>
''' Print to a network printer using the IP address. Check this when we have printer names and a DNS.
''' </summary>
''' <remarks></remarks>
Public Class PrinterIP
    Public Shared Function PrintZPL(ByVal pIP As String, ByVal psZPL As String) As Boolean
        Dim blnRet As Boolean = False
        Dim lAddress As Net.IPEndPoint
        Dim lSocket As System.Net.Sockets.Socket = Nothing
        Dim lNetStream As System.Net.Sockets.NetworkStream = Nothing
        Dim lBytes As Byte()
        Try
            lAddress = New Net.IPEndPoint(Net.IPAddress.Parse(pIP), 9100)
            lSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, _
                                 ProtocolType.Tcp)

            'the following does not block and then polls for 500 msec to see if connections succeeds.
            'blnRet is true if it does.
            'this avoids a long wait on a blocking Connect.
            lSocket.Blocking = False
            Try
                lSocket.Connect(lAddress)
            Catch ex As Exception
                ' MsgBox("Socket connection failed to : " + pIP + " with message: " + ex.Message)
            End Try
            '20181205 changed from 500 microsecs to 500,000 = 1/2 second.
            If lSocket.Poll(500000, SelectMode.SelectWrite) Then
                lSocket.Blocking = True
                lNetStream = New NetworkStream(lSocket)
                lBytes = System.Text.Encoding.ASCII.GetBytes(psZPL)
                lNetStream.Write(lBytes, 0, lBytes.Length)
                blnRet = True
            Else
                blnRet = False
                Debug.Print("Poll failed")

                '20181205 added this when debugging Apeldoorn citrix printer.
                MsgBox("Utilitities.PrintZPL: Poll of " + pIP + " failed, trying to print " + psZPL + ".")
            End If

        Catch ex As Exception 'When Not App.Debugging
            MsgBox(ex.Message & vbNewLine & ex.ToString)
        Finally
            If Not lNetStream Is Nothing Then
                lNetStream.Close()
            End If
            If Not lSocket Is Nothing Then
                lSocket.Close()
            End If
        End Try
        Return blnRet
    End Function

    'Public Shared Function PrintZPL(ByVal pIP As String, ByVal psZPL As String) As Boolean
    '    Dim blnRet As Boolean = True
    '    Dim lAddress As Net.IPEndPoint
    '    Dim lSocket As System.Net.Sockets.Socket = Nothing
    '    Dim lNetStream As System.Net.Sockets.NetworkStream = Nothing
    '    Dim lBytes As Byte()
    '    Try
    '        lAddress = New Net.IPEndPoint(Net.IPAddress.Parse(pIP), 9100)
    '        lSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, _
    '                             ProtocolType.Tcp)
    '        lSocket.Connect(lAddress)
    '        lNetStream = New NetworkStream(lSocket)
    '        lBytes = System.Text.Encoding.ASCII.GetBytes(psZPL)
    '        lNetStream.Write(lBytes, 0, lBytes.Length)
    '    Catch ex As Exception 'When Not App.Debugging
    '        blnRet = False
    '        'MsgBox(ex.Message & vbNewLine & ex.ToString)
    '    Finally
    '        If Not lNetStream Is Nothing Then
    '            lNetStream.Close()
    '        End If
    '        If Not lSocket Is Nothing Then
    '            lSocket.Close()
    '        End If
    '    End Try
    '    Return blnRet
    'End Function
End Class

''' <summary>
''' Print to a printer on a local port like USB004 using the printer name.
''' </summary>
''' <remarks></remarks>
Public Class PrinterLocal
    'print information for the spooler.
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure DOCINFO
        <MarshalAs(UnmanagedType.LPWStr)> _
        Public pDocName As String
        <MarshalAs(UnmanagedType.LPWStr)> _
        Public pOutputFile As String
        <MarshalAs(UnmanagedType.LPWStr)> _
        Public pDataType As String
    End Structure

    'interfaces in the DLL.
    <DllImport("winspool.drv", EntryPoint:="OpenPrinterW", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Public Shared Function OpenPrinter(ByVal printerName As String, ByRef hPrinter As IntPtr, ByVal printerDefaults As Integer) As Boolean
    End Function

    <DllImport("winspool.drv", EntryPoint:="ClosePrinter", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Public Shared Function ClosePrinter(ByVal hPrinter As IntPtr) As Boolean
    End Function

    <DllImport("winspool.drv", EntryPoint:="StartDocPrinterW", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Public Shared Function StartDocPrinter(ByVal hPrinter As IntPtr, ByVal level As Integer, ByRef documentInfo As DOCINFO) As Boolean
    End Function

    <DllImport("winspool.drv", EntryPoint:="EndDocPrinter", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Public Shared Function EndDocPrinter(ByVal hPrinter As IntPtr) As Boolean
    End Function

    <DllImport("winspool.drv", EntryPoint:="StartPagePrinter", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Public Shared Function StartPagePrinter(ByVal hPrinter As IntPtr) As Boolean
    End Function

    <DllImport("winspool.drv", EntryPoint:="EndPagePrinter", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Public Shared Function EndPagePrinter(ByVal hPrinter As IntPtr) As Boolean
    End Function

    <DllImport("winspool.drv", EntryPoint:="WritePrinter", SetLastError:=True, CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Public Shared Function WritePrinter(ByVal hPrinter As IntPtr, ByVal buffer As IntPtr, ByVal bufferLength As Integer, ByRef bytesWritten As Integer) As Boolean
    End Function

    Public Shared Function PrintRaw(ByVal printerName As String, ByVal origString As String) As Boolean
        Dim blnRet As Boolean = True
        Dim hPrinter As IntPtr
        Dim spoolData As New DOCINFO
        Dim dataToSend As IntPtr
        Dim dataSize As Integer
        Dim bytesWritten As Integer

        dataSize = origString.Length()
        dataToSend = Marshal.StringToCoTaskMemAnsi(origString)
        spoolData.pDocName = "OpenDrawer"
        spoolData.pDataType = "RAW"

        Try
            Call OpenPrinter(printerName, hPrinter, 0)
            Call StartDocPrinter(hPrinter, 1, spoolData)
            Call StartPagePrinter(hPrinter)
            Call WritePrinter(hPrinter, dataToSend, _
               dataSize, bytesWritten)

            EndPagePrinter(hPrinter)
            EndDocPrinter(hPrinter)
            ClosePrinter(hPrinter)
            blnRet = True
        Catch ex As Exception
            MsgBox("Error occurred: " & ex.ToString)
            blnRet = False
        Finally
            Marshal.FreeCoTaskMem(dataToSend)
        End Try
        Return blnRet
    End Function
End Class



