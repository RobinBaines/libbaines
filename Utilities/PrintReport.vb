'------------------------------------------------
'Name: PrintReport.vb.
'Function: The ReportPrintDocument will print all of the pages of a LocalReport. A ServerReport is not implemented.
'Copyright Robin Baines 2010. All rights reserved.
'Created May 2010.
'Notes: 
'Modifications:
'------------------------------------------------
Imports System
Imports System.Data
Imports System.Text
Imports System.IO
Imports System.Globalization
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Printing
Imports Microsoft.Reporting.WinForms
Imports System.Collections.Generic
Imports System.Collections.Specialized

'<summary>
'   The ReportPrintDocument will print all of the pages of a LocalReport.
'   call Print() on this class to begin printing.
'   </summary>
Public Class PrintReport
    Inherits PrintDocument
    Implements IDisposable
    Private m_pageSettings As PageSettings
    Private m_currentPage As Integer
    Private m_pages As New List(Of Stream)

    Public Sub New(localReport As LocalReport)
        Me.ReportPrintDocument(localReport, False)
        RenderAllLocalReportPages(localReport)
    End Sub

    '20201014 Allow use of other printer instead of default.
    Public Sub New(printerName As String, localReport As LocalReport, landScape As Boolean)
        If printerName.Length > 0 Then
            Me.PrinterSettings.PrinterName = printerName
        End If
        Me.ReportPrintDocument(localReport, landScape)
        RenderAllLocalReportPages(localReport)
    End Sub

    Public Sub New(serverReport As ServerReport)
        Me.ReportPrintDocument(serverReport, False)
        ' RenderAllServerReportPages(serverReport)
    End Sub

    '20201014 Added landScape.
    Private Sub ReportPrintDocument(report As Report, landScape As Boolean)
        'Set the page settings to the default defined in the report
        Dim reportPageSettings As ReportPageSettings = report.GetDefaultPageSettings()

        'The page settings object will use the default printer unless
        'PageSettings.PrinterSettings is changed.  This assumes a default printer.
        m_pageSettings = New PageSettings()
        m_pageSettings.PaperSize = reportPageSettings.PaperSize
        m_pageSettings.Margins = reportPageSettings.Margins
        If landScape = True Then
            m_pageSettings.Landscape = True
        End If
    End Sub

    Private Function CreateStream(ByVal name As String, ByVal fileNameExtension As String, ByVal encoding As Encoding, ByVal mimeType As String, ByVal willSeek As Boolean) As Stream
        Dim stream As Stream = New MemoryStream()
        m_pages.Add(stream)
        Return stream
    End Function

    Protected Overrides Sub OnBeginPrint(e As PrintEventArgs)
        MyBase.OnBeginPrint(e)
        m_currentPage = 0
    End Sub

    ' Handler for PrintPageEvents
    Protected Overrides Sub OnPrintPage(ByVal ev As PrintPageEventArgs)
        MyBase.OnPrintPage(ev)
        Dim pageToPrint As Stream = m_pages(m_currentPage)
        pageToPrint.Position = 0
        Dim pageImage As New Metafile(pageToPrint)

        ' Adjust rectangular area with printer margins.
        Dim adjustedRect As New Rectangle(ev.PageBounds.Left - CInt(ev.PageSettings.HardMarginX), _
                                          ev.PageBounds.Top - CInt(ev.PageSettings.HardMarginY), _
                                          ev.PageBounds.Width, _
                                          ev.PageBounds.Height)
        'Dim adjustedRect As New Rectangle(0, _
        '                          0, _
        '                          748, _
        '                     236)

        ' Draw a white background for the report
        ev.Graphics.FillRectangle(Brushes.White, adjustedRect)

        ' Draw the report content
        ev.Graphics.DrawImage(pageImage, adjustedRect)

        ' Prepare for the next page. Make sure we haven't hit the end.
        m_currentPage += 1
        ev.HasMorePages = (m_currentPage < m_pages.Count)
    End Sub

    Protected Overrides Sub OnQueryPageSettings(e As QueryPageSettingsEventArgs)
        e.PageSettings = m_pageSettings.Clone()
    End Sub

    Private Sub RenderAllLocalReportPages(localReport As LocalReport)
        Dim deviceInfo As String = CreateEMFDeviceInfo()

        Dim warnings As Warning()
        localReport.Render("IMAGE", deviceInfo, AddressOf LocalReportCreateStreamCallback, warnings)
    End Sub

    Private Function LocalReportCreateStreamCallback(name As String, _
        extension As String, _
        encoding As Encoding,
         mimeType As String, _
         willSeek As Boolean) As Stream
        Dim stream As MemoryStream = New MemoryStream()
        m_pages.Add(stream)
        Return stream
    End Function

    Private Function CreateEMFDeviceInfo() As String

        Dim paperSize As PaperSize = m_pageSettings.PaperSize
        Dim margins As Margins = m_pageSettings.Margins

        'The device info string defines the page range to print as well as the size of the page.
        'A start and end page of 0 means generate all pages.
        Return String.Format(
            CultureInfo.InvariantCulture,
            "<DeviceInfo><OutputFormat>emf</OutputFormat><StartPage>0</StartPage><EndPage>0</EndPage><MarginTop>{0}</MarginTop><MarginLeft>{1}</MarginLeft><MarginRight>{2}</MarginRight><MarginBottom>{3}</MarginBottom><PageHeight>{4}</PageHeight><PageWidth>{5}</PageWidth></DeviceInfo>",
            ToInches(margins.Top),
            ToInches(margins.Left),
            ToInches(margins.Right),
            ToInches(margins.Bottom),
            ToInches(paperSize.Height),
            ToInches(paperSize.Width))
    End Function

    Private Function ToInches(hundrethsOfInch As Integer) As String

        Dim inches As Double = hundrethsOfInch / 100.0
        Return inches.ToString(CultureInfo.InvariantCulture) + "in"
    End Function

    Public Overloads Sub Dispose()
        MyBase.Dispose()
        If m_pages IsNot Nothing Then
            For Each stream As Stream In m_pages
                stream.Close()
            Next
            m_pages = Nothing
        End If
        If m_pageSettings IsNot Nothing Then
            m_pageSettings = Nothing
        End If
    End Sub

End Class
