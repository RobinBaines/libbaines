'------------------------------------------------
'Name: Module ExcelInterface..vb.
'Function: 
'Copyright Robin Baines 2006. All rights reserved.
'Created July 2006.
'Notes: 
'Modifications: Feb 2007
'This interface uses specific versions of Interop.ADODB and Interop.Excel.
'Do this by browsing to the dll's in the Release directory when adding the reference.
'The properties then shows Specific Verion = true.
'This is necessary because the SaveAs fails with Access Violation read or write on 
'protected memory when using the latest versions with Excel 2000.

'20100120 RPB added check for BorderWeight.None in WriteStringToExcel. 
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data.SqlTypes
Imports Microsoft.VisualBasic.FileIO.FileSystem
Imports Microsoft.Office.Interop
Public Class ExcelInterface
    Dim xlApp As Excel.Application
    Dim oBook As Excel.Workbook
    Dim oSheet As Excel.Worksheet
    Dim oBooks As Excel.Workbooks

    'Dim iRec As Integer
    Dim strPrtArea As String
    Dim p_RowCount As Long
    Dim ColumnCount As Integer
    Dim thisThread As System.Threading.Thread
    Dim originalCulture As System.Globalization.CultureInfo
    Dim strFirstSheet As String
    Dim strLocalOutput = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    Dim strNetworkOutput = My.Computer.FileSystem.SpecialDirectories.MyDocuments    'My.Settings.OutputPath
    Public Sub New(ByVal strOutput As String)

        'Defines the network destination for files.
        strNetworkOutput = strOutput

    End Sub
    ReadOnly Property StandardPaperType() As Double
        Get
            Return Excel.XlPaperSize.xlPaperA4
        End Get
    End Property
    Public Enum ExcelStringFormats
        Bold = 1
        NotBold = 2
        BoldNotBold = 3     ''Means first non zero length string bold the rest not bold.
        NotBoldBold = 4     'Means first non zero length string not bold the rest bold.
        BoldUnderscored = 5
        Heading = 6
        Footer = 7
        NotVisible = 8
        Bold11 = 9
        Bold16 = 10
        NotBold14 = 11
        NotBold12 = 12
        Bold12 = 13
        Bold9 = 14
        NotBold8 = 15
    End Enum
    Public Enum Paths
        Local = 1
        Network = 2
    End Enum
    Public Enum BorderWeight
        None = 1
        Hairline
        Medium
        Thin
        Thick
    End Enum
    Private Function GetBorderWeight(ByVal i As BorderWeight) As Excel.XlBorderWeight
        'Convert ExcelInterface public enum to Excel value.
        If i = BorderWeight.Hairline Then Return Excel.XlBorderWeight.xlHairline
        If i = BorderWeight.Medium Then Return Excel.XlBorderWeight.xlMedium
        If i = BorderWeight.Thin Then Return Excel.XlBorderWeight.xlThin
        If i = BorderWeight.Thick Then Return Excel.XlBorderWeight.xlThick

        Return Excel.XlBorderWeight.xlHairline
    End Function

    Property RowCount() As Long
        Get
            Return p_RowCount
        End Get
        Set(ByVal value As Long)
            p_RowCount = value
        End Set
    End Property
    Private Function GetFilePath(ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal strFileName As String) As String
        Return GetDirectory(pPath, strSubDirectory) & strFileName & ".xls"
    End Function
    Private Function GetDirectory(ByVal pPath As Paths, ByVal strSubDirectory As String) As String
        Dim strPath As String
        strPath = ""
        If pPath = Paths.Local Then strPath = strLocalOutput
        If pPath = Paths.Network Then strPath = strNetworkOutput
        If Not strPath.EndsWith("\") Then strPath = strPath & "\"
        strPath = strPath & strSubDirectory
        If Not strPath.EndsWith("\") Then strPath = strPath & "\"
        Return strPath
    End Function
    Private Function MakeUniquePath(ByVal iSeed As Integer, ByVal blnUnique As Boolean, ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal strFileName As String) As String

        'Recursive call which increments the filename until the filename does not exist.
        Dim strFilePath As String
        If iSeed = 0 Then
            strFilePath = GetFilePath(pPath, strSubDirectory, strFileName)
        Else
            strFilePath = GetFilePath(pPath, strSubDirectory, strFileName & "_" & iSeed.ToString("0000"))
        End If
        If FileExists(strFilePath) = True Then
            If blnUnique = False Or iSeed > 100 Then
                My.Computer.FileSystem.DeleteFile(strFilePath)
            Else
                strFilePath = MakeUniquePath(iSeed + 1, True, pPath, strSubDirectory, strFileName)
            End If
        End If
        Return strFilePath
    End Function
    Public Sub OpenExcelBook(ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal strFileName As String, _
        ByVal blnUnique As Boolean)
        OpenExcelBook(pPath, strSubDirectory, strFileName, blnUnique, True)
    End Sub
    Public Sub OpenExcelBook(ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal strFileName As String, _
        ByVal blnUnique As Boolean, ByVal blnVisible As Boolean)

        'Ensure that thread culture setting is retained during manipulation of Excel.
        Dim strFilePath As String
        If strFileName.Length <> 0 Then

            thisThread = System.Threading.Thread.CurrentThread

            If Not My.Computer.FileSystem.DirectoryExists(GetDirectory(pPath, strSubDirectory)) Then
                My.Computer.FileSystem.CreateDirectory(GetDirectory(pPath, strSubDirectory))
            End If
            xlApp = CreateObject("Excel.Application")
            xlApp.Visible = blnVisible

            Dim strVersion = xlApp.Version


            'Solve problem caused by using US Excel on Computer with other regional setting.
            'Store current culture and set US.
            'See resetting back to original culture when session is ended in in CloseExcelBook.
            originalCulture = thisThread.CurrentCulture
            thisThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

            'Add parameter is a template and new file gets 1 added to the template name.
            'Saving also starts in the Documents directory!?
            'oBook = xlApp.Workbooks.Add("T:\Everyone\RAP\RAP2\Output\t.xls")     '(strFileName)   '"T:\Everyone\RAP\RAP2\Output\t.xls")
            oBooks = xlApp.Workbooks
            oBook = xlApp.Workbooks.Add()

            'Try to save the Excel file with the new name.
            strFilePath = MakeUniquePath(0, blnUnique, pPath, strSubDirectory, strFileName)
            Try
                oBook.SaveAs(strFilePath)
            Catch ex As Exception
                Debug.Print(ex.Message)
                'MsgBox("ERROR Path " & strFilePath & " Name " & strFileName & " " & GetDirectory(pPath, strSubDirectory))
            End Try
            'Else

            'MsgBox("Path " & strFilePath & "Name " & strFileName & " " & GetDirectory(pPath, strSubDirectory))
            'Opens with name 'Bookx'
            '   MsgBox("A file with name " & strFileName & " already exists so you will need to think up a name if you want to save.")
            'End If

            'Delete the standard sheets if necessary.
            'Dim i = xlApp.Sheets().Count
            'Do While i > 1
            'i = i - 1
            'xlApp.Sheets(i).Delete()
            'Loop

            'xlApp.Visible = False
            'xlApp.ScreenUpdating = False
            xlApp.DisplayAlerts = True
            strFirstSheet = ""
        End If
    End Sub

    Public Sub OpenExcelBook(ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal strFileName As String, ByVal strHeaderText As String, ByVal blnLandscape As Boolean)
        OpenExcelBook(pPath, strSubDirectory, strFileName, False)
        NewSheet(strHeaderText, blnLandscape, True)
    End Sub
    Public Function SetActivePrinter(ByVal strPrinter As String) As Boolean
        ', ByVal iPaperSize As Integer, ByVal iPrintQuality As Integer) As Boolean
        Dim blnRet As Boolean
        blnRet = True
        'Dim strOriginalPrinter As String
        'strOriginalPrinter = xlApp.Application.ActivePrinter
        'Dim iOriginalPrintQuality As Object
        'If iPrintQuality <> 0 Then
        'iOriginalPrintQuality = oSheet.PageSetup.PrintQuality
        'Else
        'iOriginalPrintQuality = Nothing
        'End If
        Try
            xlApp.Application.ActivePrinter = strPrinter
            '           oSheet.PageSetup.PaperSize = iPaperSize
            '          If iPrintQuality <> 0 Then oSheet.PageSetup.PrintQuality = iPrintQuality
        Catch ex As Exception
            MsgBox("Could not select Printer: " & strPrinter & " " & ex.Message, MsgBoxStyle.OkOnly)
            blnRet = False
        Finally
            'xlApp.Application.ActivePrinter = strOriginalPrinter
            'If iPrintQuality <> 0 Then
            ' oSheet.PageSetup.PrintQuality = iOriginalPrintQuality
            ' End If
        End Try
        Return blnRet
    End Function
    Public Sub NewSheet(ByVal strHeaderText As String, ByVal blnLandScape As Boolean, ByVal blnShowHeader As Boolean, _
    ByVal dLeftMargin As Double, ByVal dRightMargin As Double, ByVal dTopMargin As Double, ByVal dBottomMargin As Double, _
    ByVal dHeaderMargin As Double, _
    ByVal dFooterMargin As Double, ByVal iPaperSize As Double, ByVal blnFitToPage As Boolean, _
    ByVal blnRepeatTopRow As Boolean, ByVal iPrintQuality As Integer, ByVal iZoom As Integer)
        'If this is not the first sheet finish off the last sheet by setting the print area
        'and then add a new one.
        If Not oSheet Is Nothing Then

            'Select first cell.
            Dim mc As Object
            mc = oSheet.Cells(1, 1)
            xlApp.Sheets.Application.Range(mc.address).Activate()

            oSheet = Nothing
            oSheet = xlApp.Application.ActiveWorkbook.Sheets.Add(, xlApp.Application.ActiveSheet)

            'Make the new sheet the active one.
            oSheet.Activate()
        Else
            strFirstSheet = strHeaderText
            oSheet = xlApp.ActiveSheet
        End If

        'Give sheet a name and then set up the page.
        oSheet.Name = strHeaderText
        SetPrintArea(1, 1, 2, 2)
        PageSetup(strHeaderText, blnLandScape, blnShowHeader, dLeftMargin, dRightMargin, dTopMargin, dBottomMargin, _
            dHeaderMargin, dFooterMargin, iPaperSize, blnFitToPage, blnRepeatTopRow, iPrintQuality, iZoom)
    End Sub
    Public Sub NewSheet(ByVal strHeaderText As String, ByVal blnLandScape As Boolean, ByVal blnShowHeader As Boolean)
        NewSheet(strHeaderText, blnLandScape, blnShowHeader, _
            0.9, 0.56, 0.58, 0.58, 0, 0, StandardPaperType, True, True, 0, 0)
    End Sub
    Public Sub CloseExcelBook()

        'Set the print area based on the RowCount and Columns.
        'SetPrintArea()

        'Select all the sheets.
        'xlApp.Sheets().Select()

        'activate the first one
        If strFirstSheet.Length <> 0 Then

            'Select first cell.
            Dim mc As Object
            mc = oSheet.Cells(1, 1)
            xlApp.Sheets.Application.Range(mc.address).Activate()
            xlApp.Sheets(strFirstSheet).Activate()
        End If

        'Show the Excel application and then de couple from launching application without closing Excel.
        'Give user the control.
        xlApp.UserControl = True
        xlApp.ScreenUpdating = True
        xlApp.Visible = True
        NAR(oSheet)
        NAR(oBook)
        NAR(oBooks)

        'Not sure about this but by not calling the Yes/No/Cancel dialog is avoided.
        'xlApp.Quit()
        NAR(xlApp)
        GC.Collect()
        GC.WaitForPendingFinalizers()
        thisThread.CurrentCulture = originalCulture
    End Sub
    Public Sub CloseExcelBookAndExitExcel()


        'activate the first one
        If strFirstSheet.Length <> 0 Then

            'Select first cell.
            Dim mc As Object
            mc = oSheet.Cells(1, 1)
            xlApp.Sheets.Application.Range(mc.address).Activate()
            xlApp.Sheets(strFirstSheet).Activate()
        End If

        'Show the Excel application and then de couple from launching application without closing Excel.
        'Give user the control.
        'xlApp.UserControl = True
        'xlApp.ScreenUpdating = True
        'xlApp.Visible = True
        oBook.Save()
        xlApp.Quit()
        NAR(oSheet)
        NAR(oBook)
        NAR(oBooks)

        'Not sure about this but by not calling the Yes/No/Cancel dialog is avoided.
        'xlApp.Quit()
        NAR(xlApp)
        GC.Collect()
        GC.WaitForPendingFinalizers()
        thisThread.CurrentCulture = originalCulture
    End Sub
    Private Sub NAR(ByRef o As Object)

        'MOD RPB 21st May 2007. Change to ByRef to ensure the object variable is reset.
        'See http://support.microsoft.com/default.aspx?scid=KB;EN-US;q317109
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch
        Finally
            o = Nothing
        End Try
    End Sub
    Public Sub SetWidthandFormat(ByVal dg As DataGridView, _
    ByVal dgColumn As DataGridViewTextBoxColumn, _
        ByVal iC As Integer, ByVal strDoNotPrint As String, ByVal iWidth As Integer)

        'Set column widths and format the numerics for Excel.
        For Each iColumn As DataGridViewColumn In dg.Columns

            'Write column headers and adjust width of columns.
            Dim strFormat As String
            If CheckTag(iColumn, strDoNotPrint) Then
                iC = iC + 1

                'Set the number type in the columns.
                strFormat = ""
                If iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("N") Then
                    'strFormat = "0" & strPrecision(iColumn.DefaultCellStyle.Format)
                    strFormat = "#,##0" & strPrecision(iColumn.DefaultCellStyle.Format)
                End If
                If iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("P") Then
                    strFormat = "0" & strPrecision(iColumn.DefaultCellStyle.Format) & "%"
                End If
                If iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("C") Then
                    strFormat = "$#,##0" & strPrecision(iColumn.DefaultCellStyle.Format)
                End If
                If iColumn.Equals(dgColumn) Then
                    strFormat = "#,##0""Kg""" & strPrecision(iColumn.DefaultCellStyle.Format)
                End If
                If strFormat.Length <> 0 Then

                    'Set the format; the function aligns heading right if numeric.
                    SetColumnFormat(strFormat, 1, iC, iC)
                End If
                SetColumnWidth(iC, iWidth / 5)
            End If
        Next iColumn
    End Sub

    Public Sub SetBorders(ByVal iBorder As BorderWeight, ByVal iFirstRow As Integer, _
        ByVal iLastRow As Integer, _
        ByVal iFirstColumn As Integer, _
        ByVal iLastColumn As Integer)

        'This sets a border around the range (not in it!).
        If (iBorder <> 0) Then
            Dim mc As Object
            mc = oSheet.Cells(iFirstRow, iFirstColumn)
            Dim strRow As String
            strRow = mc.address
            mc = oSheet.Cells(iLastRow, iLastColumn)
            Dim strRow2 As String
            strRow2 = mc.address
            xlApp.Sheets.Application.Range(strRow, strRow2).Select()
            xlApp.Sheets.Application.Range(strRow, strRow2).Activate()
            With xlApp.Sheets.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = GetBorderWeight(iBorder)
                .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic 'xlAutomatic
            End With
            With xlApp.Sheets.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeTop)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = GetBorderWeight(iBorder)
                .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic 'xlAutomatic
            End With
            With xlApp.Sheets.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = GetBorderWeight(iBorder)
                .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic 'xlAutomatic
            End With
            With xlApp.Sheets.Application.Selection.Borders(Excel.XlBordersIndex.xlEdgeBottom)
                .LineStyle = Excel.XlLineStyle.xlContinuous
                .Weight = GetBorderWeight(iBorder)
                .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic 'xlAutomatic
            End With
        End If
    End Sub

    ''' <summary>
    ''' 20100120 RPB Created SwitchOffBorders. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SwitchOffBorders()
        oSheet.Cells.Select()
        oSheet.Application.ActiveWindow.DisplayGridlines = False
        oSheet.Application.Range("A1").Select()
    End Sub

    Public Sub MakeASum(ByVal iAtRow As Integer, ByVal iAtColumn As Integer, _
        ByVal iFromRow As Integer, ByVal iToRow As Integer, ByVal dDivideBy As Double)

        Dim mc As Object
        mc = oSheet.Cells(iFromRow, iAtColumn)
        Dim strFormula As String
        strFormula = "=Sum(" + mc.address
        mc = oSheet.Cells(iToRow, iAtColumn)
        If dDivideBy = 0 Or dDivideBy = 1 Then
            strFormula = strFormula + ":" + mc.address + ")"
        Else
            strFormula = strFormula + ":" + mc.address + ") / " + dDivideBy.ToString()
        End If
        xlApp.Sheets.Application.Cells(iAtRow, iAtColumn).Formula = strFormula

    End Sub
    Public Sub CreateFormula(ByVal iAtRow As Integer, ByVal iAtColumn As Integer, _
    ByVal strFormula As String)

        xlApp.Sheets.Application.Cells(iAtRow, iAtColumn).Formula = strFormula

    End Sub
    Public Sub SetColumnFormat(ByVal strFormat As String, ByVal iFirstRow As Integer, _
        ByVal iFirstColumn As Integer, _
        ByVal iLastColumn As Integer)
        SetColumnFormat(strFormat, iFirstRow, RowCount, iFirstColumn, iLastColumn)
    End Sub
    Public Sub SetColumnFormat(ByVal strFormat As String, ByVal iFirstRow As Integer, _
        ByVal iLastRow As Integer, _
        ByVal iFirstColumn As Integer, _
        ByVal iLastColumn As Integer)

        Dim mc As Object
        mc = oSheet.Cells(iFirstRow, iFirstColumn)
        Dim strRow As String
        strRow = mc.address
        mc = oSheet.Cells(iLastRow, iLastColumn)
        Dim strRow2 As String
        strRow2 = mc.address
        xlApp.Sheets.Application.Range(strRow, strRow2).Select()
        xlApp.Sheets.Application.Range(strRow, strRow2).Activate()
        If strFormat.Length <> 0 Then
            xlApp.Sheets.Application.Selection.NumberFormat = strFormat

            'RPB May 2007 Added right alighnment for numbers.
            xlApp.Sheets.Application.Selection.HorizontalAlignment = Excel.Constants.xlRight
        Else
            xlApp.Sheets.Application.Selection.WrapText = True
        End If

    End Sub
    Public Sub HideColumns(ByVal strColumns As String)
        'for example HideColumns("A:B")

        xlApp.Sheets.Application.Columns(strColumns).Select()
        xlApp.Sheets.Application.Selection.EntireColumn.Hidden = True
    End Sub
    Public Sub SetColumnBold(ByVal blnBold As Boolean, ByVal iFirstRow As Integer, ByVal iColumn As Integer)

        Dim mc As Object
        mc = oSheet.Cells(iFirstRow, iColumn)
        Dim strRow As String
        strRow = mc.address
        mc = oSheet.Cells(RowCount, iColumn)
        Dim strRow2 As String
        strRow2 = mc.address
        xlApp.Sheets.Application.Range(strRow, strRow2).Select()
        xlApp.Sheets.Application.Range(strRow, strRow2).Activate()
        xlApp.Sheets.Application.Selection.Font.Bold = blnBold
    End Sub
    Public Sub SetColumnWidth(ByVal iColumn As Integer, ByVal iWidth As Integer)

        xlApp.Sheets.Application.Columns(iColumn).ColumnWidth = iWidth

    End Sub
    Public Sub SetCellFont(ByVal fFormat As ExcelStringFormats, ByVal iRow As Integer, ByVal iCol As Integer)
        'Write Bold, NotBold or BoldNotBold. BoldNotBold means first field bold the rest not bold. 
        If fFormat = ExcelStringFormats.Bold Then
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Bold = True
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Size = 10
        ElseIf fFormat = ExcelStringFormats.NotBold Then
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Bold = False
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Size = 10
        ElseIf fFormat = ExcelStringFormats.NotBold14 Then
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Bold = False
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Size = 14
        ElseIf fFormat = ExcelStringFormats.NotBold12 Then
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Bold = False
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Size = 12
        ElseIf fFormat = ExcelStringFormats.NotBold8 Then
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Bold = False
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Size = 8
        ElseIf fFormat = ExcelStringFormats.Bold11 Then
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Bold = True
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Size = 11
        ElseIf fFormat = ExcelStringFormats.Bold16 Then
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Bold = True
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Size = 16
        ElseIf fFormat = ExcelStringFormats.Bold12 Then
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Bold = True
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Size = 12
        ElseIf fFormat = ExcelStringFormats.Bold9 Then
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Bold = True
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Size = 9
        ElseIf fFormat = ExcelStringFormats.Heading Then
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Bold = True
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Size = 16
        ElseIf fFormat = ExcelStringFormats.Footer Then
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Size = 8
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.Italic = True
        ElseIf fFormat = ExcelStringFormats.NotVisible Then
            xlApp.Sheets.Application.Cells(iRow, iCol).Font.ColorIndex = 2
        End If
    End Sub
    Public Sub WriteDataGrid(ByVal dg As DataGridView, ByVal strTagFilter As String, ByVal blnCurrentBold As Boolean, _
        ByVal iFirstColumn As Integer, ByVal blnAutofit As Boolean)

        'Display DataGridView in Excel.
        'Function requires that all headertexts of the datagridview are unique.

        'First copy DataGridView data into a recordset. 
        Dim rs As ADODB.Recordset
        rs = CreateRecordsetFromDataTable(dg, strTagFilter)

        'Write the headers
        Dim iExcelField As Integer
        iExcelField = iFirstColumn
        xlApp.Sheets.Application.Rows(RowCount).Font.Bold = True
        For Each iColumn As DataGridViewColumn In dg.Columns

            'Write column headers and adjust width of columns.
            'Align heading right if numeric
            If CheckTag(iColumn, strTagFilter) Then
                If iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("N") Or iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("P") Or _
                iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("C") Then
                    xlApp.Sheets.Application.Cells(RowCount, iExcelField + 1).HorizontalAlignment = Excel.Constants.xlRight
                End If
                xlApp.Sheets.Application.Cells(RowCount, iExcelField + 1) = iColumn.HeaderText
                If blnAutofit = True Then
                    xlApp.Sheets.Application.Columns(iExcelField + 1).ColumnWidth = iColumn.Width / 5
                End If
                iExcelField = iExcelField + 1
            End If
        Next iColumn

        'Use standard function to write the recordset to Excel.
        AdjustRowCount(1)
        Dim iFirstRow As Integer
        iFirstRow = RowCount
        Dim strFirstColumn As String
        strFirstColumn = "A"
        If iFirstColumn = 1 Then
            strFirstColumn = "B"
        End If

        Dim mc1 As Object
        mc1 = oSheet.Cells(RowCount, iFirstColumn + 1)
        'oSheet.Range(strFirstColumn & RowCount.ToString("0")).CopyFromRecordset(rs)
        oSheet.Range(mc1.address).CopyFromRecordset(rs)
        If blnCurrentBold = True And Not IsNothing(dg.CurrentRow) Then

            'Set current row in datagridview bold.
            SetRowFont(RowCount + dg.CurrentRow.Index, ExcelStringFormats.Bold)
        End If
        AdjustRowCount(dg.RowCount)

        'Adjust ColumnCount so that it includes all columns written.
        AdjustColumnCount(iExcelField)

        'Adjust the formats of the columns to match those of the data grid.
        iExcelField = 0
        Dim strFormat As String
        For Each iColumn As DataGridViewColumn In dg.Columns

            If CheckTag(iColumn, strTagFilter) Then

                'Set the number type in the columns.
                strFormat = ""
                If iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("N") Then
                    'strFormat = "0" & strPrecision(iColumn.DefaultCellStyle.Format)
                    strFormat = "#,##0" & strPrecision(iColumn.DefaultCellStyle.Format)
                End If
                If iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("P") Then
                    strFormat = "0" & strPrecision(iColumn.DefaultCellStyle.Format) & "%"
                End If
                If iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("C") Then
                    strFormat = "$#,##0" & strPrecision(iColumn.DefaultCellStyle.Format)
                End If

                'RPB Feb 2007
                SetColumnFormat(strFormat, iFirstRow, RowCount, iExcelField + 1, iExcelField + 1)
                'Dim mc As Object
                'mc = oSheet.Cells(iFirstRow, iExcelField + 1)
                'Dim strRow As String
                'strRow = mc.address
                'mc = oSheet.Cells(RowCount, iExcelField + 1)
                'Dim strRow2 As String
                'strRow2 = mc.address
                'xlApp.Sheets.Application.Range(strRow, strRow2).Select()
                'xlApp.Sheets.Application.Range(strRow, strRow2).Activate()
                'If strFormat.Length <> 0 Then
                '    xlApp.Sheets.Application.Selection.NumberFormat = strFormat
                'Else
                '    xlApp.Sheets.Application.Selection.WrapText = True
                'End If
                iExcelField = iExcelField + 1
            End If
        Next iColumn
    End Sub
    Public Sub AdjustHeightOfRow(ByVal dHeight As Double)
        xlApp.Sheets.Application.Rows(RowCount).RowHeight = dHeight
    End Sub
    Public Sub AdjustHeightOfRow(ByVal iColumn As Integer, ByVal dHeight As Double)
        xlApp.Sheets.Application.Rows(iColumn).RowHeight = dHeight
    End Sub
    Public Sub AdjustWidthOfColumn(ByVal iColumn As Integer, ByVal dWidth As Double)
        xlApp.Sheets.Application.Columns(iColumn).ColumnWidth = dWidth
    End Sub
    Public Sub SetRow(ByVal iRow As Integer)
        RowCount = iRow
    End Sub
    Public Sub SetVerticalPageBreak(ByVal iColumn As Integer)
        oSheet.Columns(iColumn).Select()
        oSheet.VPageBreaks.Add(xlApp.Sheets.Application.ActiveCell)
        'oSheet.VPageBreaks(iBreak).Location = oSheet.Range(strCoordinate)
    End Sub
    Public Sub SetHorizontalPageBreak(ByVal iRow As Integer)
        oSheet.Rows(iRow).Select()
        oSheet.HPageBreaks.Add(xlApp.Sheets.Application.ActiveCell)
        'oSheet.HPageBreaks(iBreak).Location = oSheet.Range(strCoordinate)
    End Sub
    Public Sub WriteStringToExcelAndMerge(ByVal strT As String, ByVal fFormat As ExcelStringFormats, ByVal iNrColumns As Integer)
        Dim iLastCol As Integer

        'Use to write long strings with crlf in string (remarks etc).
        iLastCol = WriteStringToExcel(strT, fFormat, 1, False, False)
        Dim mc As Object
        mc = oSheet.Cells((RowCount - 1), iLastCol - 1)
        Dim strRow As String
        strRow = mc.address
        mc = oSheet.Cells((RowCount - 1), iLastCol + iNrColumns - 1)
        Dim strRow2 As String
        strRow2 = mc.address
        xlApp.Sheets.Application.Range(strRow, strRow2).Select()
        'xlApp.Sheets.Application.Selection.Merge()
        With xlApp.Sheets.Application.Selection
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            .VerticalAlignment = Excel.XlVAlign.xlVAlignTop
            '.WrapText = Tru
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = True
            .MergeCells = True
            .Rows.AutoFit()
        End With
    End Sub
    Public Sub WriteStringToExcel(ByVal strT As String, ByVal fFormat As ExcelStringFormats)
        WriteStringToExcel(strT, fFormat, 1, False, 0)
    End Sub
    Public Sub WriteStringToExcel(ByVal strT As String, ByVal fFormat As ExcelStringFormats, _
        ByVal iBorder As BorderWeight)
        WriteStringToExcel(strT, fFormat, 1, False, iBorder)
    End Sub
    Public Sub SetRowColour(ByVal iIndex As Integer)
        'Call before Writing a string.
        xlApp.Sheets.Application.Rows(RowCount).Font.ColorIndex = iIndex
    End Sub
    Public Sub SetCellColour(ByVal iColumn As Integer, ByVal iIndex As Integer)

        Dim mc As Object
        mc = oSheet.Cells(RowCount, iColumn)
        Dim strRow As String
        strRow = mc.address
        'mc = oSheet.Cells(RowCount, iColumn)
        'Dim strRow2 As String
        'strRow2 = mc.address
        xlApp.Sheets.Application.Range(strRow).Select()
        xlApp.Sheets.Application.Range(strRow).Activate()
        xlApp.Sheets.Application.Selection.Interior.ColorIndex = iIndex
        xlApp.Sheets.Application.Selection.Interior.Pattern = Excel.XlPattern.xlPatternSolid
    End Sub

    Public Function WriteStringToExcel(ByVal strT As String, _
    ByVal fFormat As ExcelStringFormats, _
    ByVal iStartColumn As Integer, _
    ByVal blnRightAlign As Boolean, _
    ByVal iBorder As BorderWeight) As Integer

        'Write the string to Excel using a separate column for each field in the string.
        'Field separator is #.
        Dim strSplit() As String
        Dim s As String
        Dim iCol As Integer

        'Write Bold, NotBold or BoldNotBold. BoldNotBold means first field bold the rest not bold. 
        If fFormat = ExcelStringFormats.Bold Then
            xlApp.Sheets.Application.Rows(RowCount).Font.Bold = True
            xlApp.Sheets.Application.Rows(RowCount).Font.Size = 10
        ElseIf fFormat = ExcelStringFormats.NotBold Then
            xlApp.Sheets.Application.Rows(RowCount).Font.Bold = False
            xlApp.Sheets.Application.Rows(RowCount).Font.Size = 10
        ElseIf fFormat = ExcelStringFormats.NotBold14 Then
            xlApp.Sheets.Application.Rows(RowCount).Font.Bold = False
            xlApp.Sheets.Application.Rows(RowCount).Font.Size = 14
        ElseIf fFormat = ExcelStringFormats.NotBold12 Then
            xlApp.Sheets.Application.Rows(RowCount).Font.Bold = False
            xlApp.Sheets.Application.Rows(RowCount).Font.Size = 12
        ElseIf fFormat = ExcelStringFormats.NotBold8 Then
            xlApp.Sheets.Application.Rows(RowCount).Font.Bold = False
            xlApp.Sheets.Application.Rows(RowCount).Font.Size = 8
        ElseIf fFormat = ExcelStringFormats.Bold11 Then
            xlApp.Sheets.Application.Rows(RowCount).Font.Bold = True
            xlApp.Sheets.Application.Rows(RowCount).Font.Size = 11
        ElseIf fFormat = ExcelStringFormats.Bold16 Then
            xlApp.Sheets.Application.Rows(RowCount).Font.Bold = True
            xlApp.Sheets.Application.Rows(RowCount).Font.Size = 16
        ElseIf fFormat = ExcelStringFormats.Bold12 Then
            xlApp.Sheets.Application.Rows(RowCount).Font.Bold = True
            xlApp.Sheets.Application.Rows(RowCount).Font.Size = 12
        ElseIf fFormat = ExcelStringFormats.Bold9 Then
            xlApp.Sheets.Application.Rows(RowCount).Font.Bold = True
            xlApp.Sheets.Application.Rows(RowCount).Font.Size = 9
        ElseIf fFormat = ExcelStringFormats.Heading Then
            xlApp.Sheets.Application.Rows(RowCount).Font.Bold = True
            xlApp.Sheets.Application.Rows(RowCount).Font.Size = 16
        ElseIf fFormat = ExcelStringFormats.Footer Then
            xlApp.Sheets.Application.Rows(RowCount).Font.Size = 8
            xlApp.Sheets.Application.Rows(RowCount).Font.Italic = True
        ElseIf fFormat = ExcelStringFormats.NotVisible Then
            xlApp.Sheets.Application.Rows(RowCount).Font.ColorIndex = 2
        End If
        If iBorder = BorderWeight.None Then

        Else
            '20100120 RPB added check for BorderWeight.None. 
            If (iBorder <> 0 And iBorder <> BorderWeight.None) Then
                'With xlApp.Sheets.Application.Rows(RowCount).Borders(Excel.XlBordersIndex.xlEdgeRight)
                '    .LineStyle = Excel.XlLineStyle.xlContinuous
                '    .Weight = GetBorderWeight(iBorder)
                '    .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic 'xlAutomatic
                'End With
                With xlApp.Sheets.Application.Rows(RowCount).Borders(Excel.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = GetBorderWeight(iBorder)
                    .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic 'xlAutomatic
                End With
                'With xlApp.Sheets.Application.Rows(RowCount).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                '    .LineStyle = Excel.XlLineStyle.xlContinuous
                '    .Weight = GetBorderWeight(iBorder)
                '    .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic 'xlAutomatic
                'End With
                With xlApp.Sheets.Application.Rows(RowCount).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = GetBorderWeight(iBorder)
                    .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic 'xlAutomatic
                End With
            End If
        End If

        '        ElseIf fFormat = ExcelStringFormats.BoldUnderscored Then
        '        xlApp.Sheets.Application.Rows(RowCount).Font.Bold = True
        '        xlApp.Sheets.Application.Rows(RowCount).Font.Underline = 2  'Excel.LineFormat.xlUnderlineStyleSingle
        '        xlApp.Sheets.Application.Rows(RowCount).Font.Size = 10

        strSplit = strT.Split(New [Char]() {"#"c})
        iCol = iStartColumn

        Dim blnBoldNotBold = False  'used to flag first string when using BoldNotBold.
        For Each s In strSplit
            If fFormat = ExcelStringFormats.BoldNotBold Then

                'Means first non zero length string bold the rest not bold.
                If blnBoldNotBold = False And s.Length <> 0 Then
                    blnBoldNotBold = True
                    xlApp.Sheets.Application.Cells(RowCount, iCol).Font.Bold = True
                Else
                    xlApp.Sheets.Application.Cells(RowCount, iCol).Font.Bold = False
                End If
            End If

            '20100120 RPB added check for BorderWeight.None. 
            If (iBorder <> 0 And iBorder <> BorderWeight.None) Then
                With xlApp.Sheets.Application.Cells(RowCount, iCol).Borders(Excel.XlBordersIndex.xlEdgeRight)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = GetBorderWeight(iBorder)
                    .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic 'xlAutomatic
                End With
                'With xlApp.Sheets.Application.Cells(RowCount, iCol).Borders(Excel.XlBordersIndex.xlEdgeTop)
                '    .LineStyle = Excel.XlLineStyle.xlContinuous
                '    .Weight = GetBorderWeight(iBorder)
                '    .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic 'xlAutomatic
                'End With
                With xlApp.Sheets.Application.Cells(RowCount, iCol).Borders(Excel.XlBordersIndex.xlEdgeLeft)
                    .LineStyle = Excel.XlLineStyle.xlContinuous
                    .Weight = GetBorderWeight(iBorder)
                    .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic 'xlAutomatic
                End With
                'With xlApp.Sheets.Application.Cells(RowCount, iCol).Borders(Excel.XlBordersIndex.xlEdgeBottom)
                '    .LineStyle = Excel.XlLineStyle.xlContinuous
                '    .Weight = GetBorderWeight(iBorder)
                '    .ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic 'xlAutomatic
                'End With
            End If
            If fFormat = ExcelStringFormats.NotBoldBold Then

                'Means first non zero length string not bold the rest bold.
                If blnBoldNotBold = False And s.Length <> 0 Then
                    blnBoldNotBold = True
                    xlApp.Sheets.Application.Cells(RowCount, iCol).Font.Bold = False
                Else
                    xlApp.Sheets.Application.Cells(RowCount, iCol).Font.Bold = True
                End If
            End If

            If fFormat = ExcelStringFormats.BoldUnderscored Then

                'Means first non zero length string not bold the rest bold.
                xlApp.Sheets.Application.Rows(RowCount).Font.Bold = True
                xlApp.Sheets.Application.Rows(RowCount).Font.Size = 10
                If blnBoldNotBold = False And s.Length <> 0 Then
                    blnBoldNotBold = True
                    xlApp.Sheets.Application.Cells(RowCount, iCol).Font.Underline = 2  'Excel.LineFormat.xlUnderlineStyleSingle
                Else
                    xlApp.Sheets.Application.Cells(RowCount, iCol).Font.Underline = 1
                End If
            End If
            If blnRightAlign = True Then
                xlApp.Sheets.Application.Cells(RowCount, iCol).HorizontalAlignment = Excel.Constants.xlRight
            End If


            'xlApp.Sheets.Application.Cells(RowCount, iCol) = "'" & s
            xlApp.Sheets.Application.Cells(RowCount, iCol) = s
            iCol = iCol + 1
        Next s
        AdjustColumnCount(iCol)
        AdjustRowCount(1)
        Return iCol
    End Function
    Public Sub strCreateExcelSheet(ByVal dg As DataGridView, ByVal strFileName As String, ByVal strHeaderText As String, _
    ByVal strTagFilter As String)

        'Opens excel, writes datagridview and closes.
        OpenExcelBook(Paths.Local, "", strFileName, strHeaderText, True)

        'The empty string means all columns are printed without filtering.
        WriteDataGrid(dg, strTagFilter, False, 0, True)
        SetPrintArea()
        CloseExcelBook()
    End Sub
    Public Sub SetPrintArea()
        Dim mc As Object
        mc = oSheet.Cells(1, 1)
        strPrtArea = mc.Address()

        'RPB Feb 2007 Decrement ColumnCount before using to get exactly right.
        mc = oSheet.Cells(RowCount, ColumnCount - 1)
        strPrtArea = strPrtArea & ":" & mc.Address()
        oSheet.PageSetup.PrintArea = strPrtArea
    End Sub
    Public Sub SetPrintArea(ByVal iFirstRow As Integer, ByVal iFirstColumn As Integer, _
    ByVal iLastRow As Integer, ByVal iLastColumn As Integer)
        Dim mc As Object
        mc = oSheet.Cells(iFirstRow, iFirstColumn)
        strPrtArea = mc.Address()
        mc = oSheet.Cells(iLastRow, iLastColumn)
        strPrtArea = strPrtArea & ":" & mc.Address()
        oSheet.PageSetup.PrintArea = strPrtArea
    End Sub
    Public Sub SetRowFont(ByVal rCount As Long, ByVal fFormat As ExcelStringFormats)
        If fFormat = ExcelStringFormats.Bold Then
            xlApp.Sheets.Application.Rows(rCount).Font.Bold = True
        End If
        If fFormat = ExcelStringFormats.NotBold Then
            xlApp.Sheets.Application.Rows(rCount).Font.Bold = False
        End If
    End Sub
    Public Sub PrintMasterChild(ByVal dgMaster As DataGridView, ByVal dgChild As DataGridView, _
    ByVal strSubDirectory As String, ByVal strFileName As String, ByVal strHeader As String, ByVal strTagFilter As String)
        OpenExcelBook(Paths.Local, strSubDirectory, strFileName, False)
        NewSheet(strHeader, True, True)
        WriteDataGrid(dgMaster, strTagFilter, True, 0, False)
        WriteStringToExcel(" ", True, False)
        WriteDataGrid(dgChild, strTagFilter, False, 0, False)
        WriteStringToExcel(" ", True, False)
        SetAutofit()
        SetPrintArea()
        CloseExcelBook()
    End Sub
    Public Sub InsertPicture(ByVal strImageName As String, ByVal strAtCell As String)


        'xlApp.ActiveSheet.Pictures.Insert(strImageName) '"T:\Everyone\RAP\RAP2\Output\JPGs\image002.jpg"). _
        Dim p As Object = xlApp.Sheets.Application.ActiveSheet.Pictures.Insert(strImageName) '"T:\Everyone\RAP\RAP2\Output\JPGs\image002.jpg"). _
        p.left = xlApp.Sheets.Application.ActiveSheet.Range(strAtCell).left
        p.top = xlApp.Sheets.Application.ActiveSheet.Range(strAtCell).top
        'Excel 2003 could use selected cell to position the picture but
        '2007 needs an explicit pixel position.
        'p.left = 1
        'p.top = 5

        p = Nothing


        'xlApp.Sheets.Application.ActiveSheet.Range(strAtCell).Select()


    End Sub
    Public Sub SetBorder()
        Dim strRow As String
        strRow = "A" & (RowCount - 1).ToString("0")
        Dim mc As Object
        mc = oSheet.Cells((RowCount - 1), ColumnCount)
        Dim strRow2 As String
        strRow2 = mc.address
        With oSheet.Range(strRow, strRow2).Borders(Excel.XlBordersIndex.xlEdgeBottom)   'Selection.Borders(xlEdgeBottom)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            '.Weight = Excel.Constants.xlt.xlThin
            .ColorIndex = Excel.Constants.xlAutomatic
        End With
    End Sub
    Public Sub SetAutofit()

        'Use Row and Column counts to select an area and autofit to it.
        Dim str1 As String
        str1 = "A:" 'To limit to the last row do the following. & (RowCount - 1).ToString("0")
        Dim mc As Excel.Range
        Dim rc As Long
        rc = RowCount - 1
        If rc > 40 Then
            rc = 40
        End If

        mc = oSheet.Cells(rc, (ColumnCount - 1))
        Dim str2 As String
        str2 = mc.Address

        'Returns an address like $G$2 so get the column.
        If str2.StartsWith("$") Then
            str1 = str1 & str2.Substring(1, 1)
        Else
            str1 = str1 & str2.Substring(0, 1)
        End If

        'Create a column range like A:H
        oSheet.Columns(str1).Select()
        xlApp.Sheets.Application.Selection.Columns.AutoFit()
    End Sub
    'PRIVATE Subs and functions.
    'Is also in Utilities!!!
    Private Function CheckTag(ByVal iColumn As DataGridViewColumn, ByVal strTagFilter As String) As Boolean
        If IsNothing(iColumn.Tag) Then Return True
        If iColumn.Tag.ToString.Length = 0 Then
            Return True
        End If
        If strTagFilter.Contains(iColumn.Tag) Then Return False
        Return True
    End Function
    Private Sub AdjustColumnCount(ByVal iCols As Integer)
        If ColumnCount < iCols Then
            ColumnCount = iCols
        End If
    End Sub
    Private Sub AdjustRowCount(ByVal iAddRows As Integer)
        RowCount = RowCount + iAddRows
    End Sub
    Private Function CreateRecordsetFromDataTable(ByVal dg As DataGridView, ByVal strTagFilter As String) As ADODB.Recordset

        'Create columns in an ADODB.Recordset from the columns in the DataGridView.
        Dim rs As New ADODB.Recordset
        Dim FieldAttr As ADODB.FieldAttributeEnum
        FieldAttr = ADODB.FieldAttributeEnum.adFldIsNullable Or _
            ADODB.FieldAttributeEnum.adFldIsNullable Or ADODB.FieldAttributeEnum.adFldUpdatable
        For Each iColumn As DataGridViewColumn In dg.Columns
            If CheckTag(iColumn, strTagFilter) = True Then
                Dim FieldType As ADODB.DataTypeEnum
                FieldType = ADODB.DataTypeEnum.adVarWChar
                If iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("N") Or iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("P") Or _
                iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("C") Then
                    If iColumn.DefaultCellStyle.Format.Contains("0") Then
                        FieldType = ADODB.DataTypeEnum.adInteger
                    Else
                        FieldType = ADODB.DataTypeEnum.adDouble
                    End If
                End If
                If FieldType = ADODB.DataTypeEnum.adVarWChar Then
                    rs.Fields.Append(iColumn.HeaderText, FieldType, 4000)
                Else
                    rs.Fields.Append(iColumn.HeaderText, FieldType)
                End If
                rs.Fields(iColumn.HeaderText).Attributes = FieldAttr
            End If
        Next
        rs.Open()

        'Copies all the row of the DataGridView into the recordset
        For Each iRow As DataGridViewRow In dg.Rows
            'Try to fish out the last row which is just a row of 0's.
            If iRow.IsNewRow = False Then
                rs.AddNew()
                Dim iRecordSetColumn As Integer
                iRecordSetColumn = 0
                Dim iDataGridColumn As Integer
                iDataGridColumn = 0
                For Each iColumn As DataGridViewColumn In dg.Columns
                    If CheckTag(iColumn, strTagFilter) = True Then
                        If Not iRow.Cells(iDataGridColumn).Value Is System.DBNull.Value Then
                            rs(iRecordSetColumn).Value = iRow.Cells(iDataGridColumn).Value
                        Else
                            If rs(iRecordSetColumn).Type = ADODB.DataTypeEnum.adInteger Or _
                                rs(iRecordSetColumn).Type = ADODB.DataTypeEnum.adDouble Then
                                rs(iRecordSetColumn).Value = 0
                            Else
                                rs(iRecordSetColumn).Value = ""
                            End If
                        End If
                        iRecordSetColumn = iRecordSetColumn + 1
                    End If
                    iDataGridColumn = iDataGridColumn + 1
                Next
            End If
        Next

        'Moves to the first record in recordset
        If Not rs.BOF Then rs.MoveFirst()
        Return rs

    End Function
    Public Sub SetFitTo1Page()
        oSheet.PageSetup.FitToPagesWide = 1
        oSheet.PageSetup.FitToPagesTall = 1
    End Sub
    Private Function PageSetup(ByVal strHeaderText As String, ByVal blnLandscape As Boolean, ByVal blnShowHeader As Boolean, _
    ByVal dLeftMargin As Double, ByVal dRightMargin As Double, ByVal dTopMargin As Double, ByVal dBottomMargin As Double, _
    ByVal dHeaderMargin As Double, _
    ByVal dFooterMargin As Double, ByVal iPaperSize As Integer, ByVal blnFitToPage As Boolean, _
    ByVal blnRepeatTopRow As Boolean, ByVal iPrintQuality As Integer, ByVal iZoom As Integer) As Boolean

        Dim blnRet As Boolean
        'Adjust the Excel page setup for landscape, displaying the HeaderText in the header of the print out. 
        Try
            blnRet = True

            With oSheet.PageSetup
                If blnRepeatTopRow = True Then
                    .PrintTitleRows = "$1:$1"
                Else
                    .PrintTitleRows = ""
                End If

                .PrintTitleColumns = ""
                .LeftMargin = xlApp.Sheets.Application.InchesToPoints(dLeftMargin)
                .RightMargin = xlApp.Sheets.Application.InchesToPoints(dRightMargin)
                .TopMargin = xlApp.Sheets.Application.InchesToPoints(dTopMargin)
                .BottomMargin = xlApp.Sheets.Application.InchesToPoints(dBottomMargin)
                .HeaderMargin = xlApp.Sheets.Application.InchesToPoints(dHeaderMargin)
                .FooterMargin = xlApp.Sheets.Application.InchesToPoints(dFooterMargin)
                If blnShowHeader = True Then
                    .LeftHeader = "Using data from " & strHeaderText & "."
                    .CenterHeader = "Printed on &D &T. "
                    .RightHeader = "Page &P of &N"
                    .LeftFooter = ""
                    .CenterFooter = ""
                    .RightFooter = "File created by " & System.Windows.Forms.Application.ProductName & "."
                Else
                    .LeftHeader = ""
                    .CenterHeader = ""
                    .RightHeader = ""
                    .LeftFooter = ""
                    .CenterFooter = ""
                    .RightFooter = ""
                End If
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
                .CenterHorizontally = False
                .CenterVertically = False
                If blnLandscape = True Then
                    .Orientation = Excel.XlPageOrientation.xlLandscape
                Else
                    .Orientation = Excel.XlPageOrientation.xlPortrait
                End If
                .Draft = False
                If iPaperSize <> 0 Then
                    .PaperSize = iPaperSize     'Excel.XlPaperSize.xlPaperA4
                End If
                If iPrintQuality <> 0 Then
                    .PrintQuality = iPrintQuality
                End If

                '.FirstPageNumber = xlAutomatic
                .Order = Excel.XlOrder.xlDownThenOver
                .BlackAndWhite = False
                If blnFitToPage = True Then
                    .Zoom = 70
                    .Zoom = False   'activate the Fit parameters.
                    .FitToPagesWide = 1
                    .FitToPagesTall = False
                End If
                If iZoom = 0 Then
                    .Zoom = False
                Else
                    .Zoom = iZoom
                End If
                '            .PrintArea = strPrtArea '"R1C1:R26C184"
            End With
        Catch ex As Exception
            MsgBox("Could not set up the page. Check whether printer was selected. " & ex.Message, MsgBoxStyle.OkOnly)
            blnRet = False
        Finally
            ColumnCount = 1
            RowCount = 1
        End Try
        Return blnRet
    End Function
    Public Function strPrecision(ByVal strFormat As String) As String

        'Return a string of '0' equal in length to the number in the second position in strFormat.
        'N4 returns 0000
        'C2 returns 00
        Dim str As String
        str = ""
        Dim iPrecision = CType(strFormat.Substring(1), Integer)
        If iPrecision > 0 Then
            str = str & "."
            Do While iPrecision > 0
                str = str & "0"
                iPrecision = iPrecision - 1
            Loop
        End If
        Return str
    End Function
End Class
