'------------------------------------------------
'Name: Module CXML.vb.
'Function: 
'Copyright Robin Baines 2024. All rights reserved.
'Created Oct 2024.
'Notes: 
'Modifications: 
'------------------------------------------------
Imports ClosedXML.Excel
Imports System.Globalization
Imports Microsoft.VisualBasic.FileIO.FileSystem
Imports System.Windows.Forms
''' <summary>
''' 
''' </summary>
''' <remarks></remarks>

Public Class CXML
    Dim strLocalOutput = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    Dim strNetworkOutput = My.Computer.FileSystem.SpecialDirectories.MyDocuments    'My.Settings.OutputPath
    Dim strFileName As String
    Dim strOriginalFileName As String
    Dim blnLandscape As Boolean 'Store parameter between Open and closing a sheet.

    Dim workbook As XLWorkbook
    Dim worksheet As IXLWorksheet
    Dim iRow As Integer = 1

    ''' <summary>
    ''' keep track of the number of rows in a grid.
    ''' </summary>
    Dim p_RowCount As Long
    Property RowCount() As Long
        Get
            Return p_RowCount
        End Get
        Set(ByVal value As Long)
            p_RowCount = value
        End Set
    End Property

    Dim ColumnCount As Integer

    Public Sub New(ByVal strOutput As String)

        'Defines the network destination for files.
        strNetworkOutput = strOutput
        Dim value As NumberFormatInfo
        value = NumberFormatInfo.CurrentInfo
        iRow = 1
    End Sub
    Public Function GetDirectory(ByVal pPath As Paths, ByVal strSubDirectory As String) As String

        'Get the directory using the local/network specifier.
        Dim strPath As String
        strPath = ""
        If pPath = Paths.Local Then strPath = strLocalOutput
        If pPath = Paths.Network Then strPath = strNetworkOutput
        If Not strPath.EndsWith("\") Then strPath = strPath & "\"
        strPath = strPath & strSubDirectory
        If Not strPath.EndsWith("\") Then strPath = strPath & "\"
        Return strPath
    End Function

    Private Function GetFilePath(ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal strFileName As String) As String

        'Create path to saved book.
        Return GetDirectory(pPath, strSubDirectory) & strFileName & ".xlsx"
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

    Private Sub AdjustColumnCount(ByVal iCols As Integer)
        If ColumnCount < iCols Then
            ColumnCount = iCols
        End If
    End Sub

    Private Sub AdjustRowCount(ByVal iAddRows As Integer)
        RowCount = RowCount + iAddRows
    End Sub
    Public Function OpenExcelBook(ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal _strFileName As String,
   ByVal blnUnique As Boolean) As Boolean
        Dim blnFailed As Boolean = False
        strFileName = _strFileName
        workbook = Nothing
        If strFileName.Length <> 0 Then

            Try 'GO899 No permission to write file should be visible for endusers.

                'Create target directory if it does not exist already.
                If Not My.Computer.FileSystem.DirectoryExists(GetDirectory(pPath, strSubDirectory)) Then

                    'Create target directory if it does not exist already.
                    My.Computer.FileSystem.CreateDirectory(GetDirectory(pPath, strSubDirectory))
                End If

                Try

                    'Try to save the Excel file with the new name.
                    strFileName = MakeUniquePath(0, blnUnique, pPath, strSubDirectory, _strFileName)
                    strOriginalFileName = strFileName
                    workbook = New XLWorkbook()
                Catch ex As Exception
                    blnFailed = True
                    MsgBox("CXML.OpenExcelBook Could not Create File  " + GetFilePath(pPath, strSubDirectory, strFileName) + " in OpenExcelBook. Check r/w permission. " + ex.Message)
                End Try
            Catch ex As Exception
                blnFailed = True
                MsgBox("CXML.OpenExcelBook Could not CreateDirectory  " + GetDirectory(pPath, strSubDirectory) + " in OpenExcelBook. Check r/w permission. " + ex.Message)
            End Try
        End If
        Return blnFailed
    End Function

    Public Sub NewSheet(ByVal _strSheetName As String,
                        ByVal _strHeaderText As String,
                        ByVal _blnLandScape As Boolean,
                        ByVal _strFooter As String)
        If workbook IsNot Nothing Then
            worksheet = workbook.Worksheets.Add(_strSheetName)
        End If
    End Sub

    Public Sub CloseExcelBook(blnOpen As Boolean)
        If workbook IsNot Nothing Then
            workbook.SaveAs(strFileName)
            If blnOpen Then
                Process.Start(New ProcessStartInfo(strFileName))
            End If
        End If
    End Sub


    ''' <summary>
    ''' It was possible to prevent printing a datagrid by setting the 
    ''' </summary>
    ''' <param name="iColumn"></param>
    ''' <param name="strTagFilter"></param>
    ''' <returns></returns>
    Private Function CheckTag(ByVal iColumn As DataGridViewColumn, ByVal strTagFilter As String) As Boolean
        If IsNothing(iColumn.Tag) Then Return True
        If iColumn.Tag.ToString.Length = 0 Then
            Return True
        End If
        If strTagFilter.Contains(iColumn.Tag) Then Return False
        Return True
    End Function

    Public Sub WriteDataGrid(ByVal dg As DataGridView, ByVal strTagFilter As String, ByVal blnCurrentBold As Boolean,
        ByVal iFirstColumn As Integer, ByVal blnAutofit As Boolean)
        WriteDataGrid(dg, strTagFilter, blnCurrentBold, iFirstColumn, blnAutofit, True, True)
    End Sub

    'Dim strDefaultType As String = "String"
    'Private Function pWriteCell(ByVal strValue As String, ByVal strStyle As String, ByVal strType As String) As String

    '    If strType = "DateTime" Then
    '        If strValue.Length = 0 Then
    '            strType = strDefaultType
    '        Else
    '            Dim dt As DateTime
    '            dt = strValue
    '            strValue = dt.Year & "-" & dt.Month.ToString("0#") & "-" & dt.Day.ToString("0#") & "T00:00:00.000"
    '        End If
    '    End If

    '    'Check whether we really do have a number. If not change to String.
    '    '20110602 RPB Modified pWriteCell. Check whether a number is really a number before replacing , by . because System.Convert.ToDouble is sensitive to regional setting.
    '    If strType = "Number" Then
    '        Try
    '            '20110504
    '            If strValue.Length > 0 Then
    '                Dim dD As Double = System.Convert.ToDouble(strValue)
    '            End If
    '        Catch ex As Exception
    '            strType = strDefaultType
    '            strStyle = "NormalRightAligned"
    '        End Try
    '    End If

    '    'See comment on regional settings above.
    '    '20081105 Added this after testing on US Excel.
    '    If strType = "Number" And strSeparator = "," Then
    '        'originalCulture = System.Threading.Thread.CurrentThread.CurrentCulture
    '        strValue = strValue.Replace(".", "#")
    '        strValue = strValue.Replace(",", ".")
    '        strValue = strValue.Replace("#", ",")
    '    End If

    '    '20081112 Filter @, < and >.
    '    'strValue = FilterXML(strValue)
    '    'Dim strVal As String
    '    'If strStyle <> "" Then
    '    '    strVal = "<Cell ss:StyleID='" & strStyle & "'>"
    '    'Else
    '    '    strVal = "<Cell>"
    '    'End If

    '    '20111118 Users can alter value for True and False.
    '    'If strValue = "True" Then strValue = _strTrue
    '    'If strValue = "False" Then strValue = _strFalse

    '    'If strValue.Length <> 0 Then
    '    '    strVal = strVal & "<Data ss:Type='" & strType & "'>" & strValue & "</Data>"
    '    'End If
    '    'strVal = strVal + "<NamedCell ss:Name='Print_Area'/></Cell>"
    '    'StoreXMLData(strVal)

    '    'worksheet.Cell(iRow, iCol).Value = c.strValue
    '    'worksheet.Cell(iRow, iCol).Style.Font.Bold = (c.blnBold And c.iStyle = ExcelStringFormats.NoStyle)


    '    'stW.WriteLine(strVal)

    '    Return strValue
    'End Function

    Private Function strPrecision(ByVal strFormat As String) As String

        'Return a string of '0' equal in length to the number in the second position in strFormat.
        'N4 returns 0000
        'C2 returns 00
        Dim str As String
        str = ""
        Dim iPrecision As Integer = 0
        If Not Integer.TryParse(strFormat, iPrecision) Then
        Else
            If iPrecision > 0 Then
                str = "0."
                Do While iPrecision > 0
                    str = str & "0"
                    iPrecision = iPrecision - 1
                Loop
            End If
        End If
        Return str
    End Function

    Public Function WriteStringToExcel(ByVal Cells As List(Of XMLExcelCell),
        ByVal iStartColumn As Integer,
        ByVal blnRightAlign As Boolean, ByVal iStyle As ExcelStringFormats,
                                       blnHeader As Boolean) As Integer




        'Write the string to Excel using a separate column for each field in the string.
        Dim iCol As Integer
            iCol = iStartColumn
            Dim c As XMLExcelCell
        'Dim strStyle As String
        'Dim strType As String
        If workbook IsNot Nothing Then
            For Each c In Cells

                'Mod RPB June 2008. Dg could contain nulls. So replace with "" to avoid gaps.
                If c.strValue Is Nothing Then
                    c.strValue = ""
                End If

                'If c.blnBold = True And c.iStyle = ExcelStringFormats.NoStyle Then
                '    strStyle = strGetStyle(ExcelStringFormats.Bold10)
                'Else
                '    strStyle = strGetStyle(c.iStyle)
                'End If
                'strType = strGetType(c.iStyle)

                If blnHeader = False Then
                    If Not worksheet.Cell(iRow, iCol).Style Is Nothing Then
                        Dim sNFId As String = ""
                        If c.NumberFormat.Length > 0 Then
                            If c.NumberFormat.Substring(0, 1) = "N" Then
                                If c.NumberFormat.Substring(1, 1) = "0" Then
                                    worksheet.Cell(iRow, iCol).Style.NumberFormat.NumberFormatId = 1
                                Else
                                    'N1 - N6 => 0.0 - 0.000000
                                    worksheet.Cell(iRow, iCol).Style.NumberFormat.Format = strPrecision(c.NumberFormat.Substring(1, 1))
                                End If
                            Else
                                If c.NumberFormat.Substring(0, 2) = "\€" Then ' for example \€ 0.0000
                                    worksheet.Cell(iRow, iCol).Style.NumberFormat.Format = "€" + c.NumberFormat.Substring(2)
                                Else
                                    If c.NumberFormat.Contains("%") Then 'for example 0.000\%
                                        worksheet.Cell(iRow, iCol).Style.NumberFormat.Format = c.NumberFormat
                                    End If
                                End If
                            End If
                        End If
                        Try
                            Dim theValue As Decimal = c.strValue
                            worksheet.Cell(iRow, iCol).Value = theValue
                        Catch ex As Exception
                            worksheet.Cell(iRow, iCol).Style.NumberFormat.NumberFormatId = 0
                            worksheet.Cell(iRow, iCol).Value = c.strValue
                        End Try
                    End If
                Else

                    worksheet.Cell(iRow, iCol).Style.NumberFormat.NumberFormatId = 0 'is the general format code for Excel.
                    worksheet.Cell(iRow, iCol).Value = c.strValue
                End If

                If Not worksheet.Cell(iRow, iCol).Style.Font Is Nothing Then
                    worksheet.Cell(iRow, iCol).Style.Font.Bold = (c.blnBold And c.iStyle = ExcelStringFormats.NoStyle)
                End If


                iCol = iCol + 1

            Next c

            Try
                If blnHeader Then
                    Dim rngTable = worksheet.Range(iRow, iStartColumn, iRow, iCol)
                    'Formatting headers
                    Dim rngHeaders = rngTable.Range(iRow, iStartColumn, iRow, iCol)
                    If blnRightAlign Then
                        rngHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right
                    Else
                        rngHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
                    End If
                    rngHeaders.Style.Font.Bold = True
                End If
            Catch ex As Exception
                MsgBox("Error writing Excel header. " + ex.Message)
            End Try
            iRow += 1

            AdjustColumnCount(iCol)
            AdjustRowCount(1)
        End If

        Return iCol
    End Function

    Public Sub WriteDataGrid(ByVal dg As DataGridView, ByVal strTagFilter As String, ByVal blnCurrentBold As Boolean,
        ByVal iFirstColumn As Integer, ByVal blnAutofit As Boolean, ByVal blnWriteColumnWidths As Boolean,
        ByVal blnWriteHeaders As Boolean)

        'Display DataGridView in Excel.
        'All headertexts of the datagridview must be unique.
        Dim Cells As List(Of XMLExcelCell)
        Dim ColumnWidth As List(Of Int32)
        Cells = New List(Of XMLExcelCell)(dg.Columns.Count)
        Cells.Clear()
        ColumnWidth = New List(Of Int32)(dg.Columns.Count)
        ColumnWidth.Clear()
        Dim iCellsColumn = 0

        'Set up column widths, formats and headers.
        For Each iColumn As DataGridViewColumn In dg.Columns

            'returns false if DONOTPRINT is in the Tag property of a DataGridView Column.
            If CheckTag(iColumn, strTagFilter) = True Then

                'Decimal values are aligned right.
                If iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("N") Or iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("P") _
                    Or iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("C") Then
                    Cells.Add(New XMLExcelCell("", ExcelStringFormats.Bold10RightAligned, iColumn.DefaultCellStyle.Format))
                Else
                    Cells.Add(New XMLExcelCell("", ExcelStringFormats.Bold10, iColumn.DefaultCellStyle.Format))
                End If
                Cells(iCellsColumn).strValue = iColumn.HeaderText()
                ColumnWidth.Add(iColumn.Width)
                iCellsColumn = iCellsColumn + 1
            End If
        Next

        If blnWriteHeaders = True Then
            WriteStringToExcel(Cells, 1, False, ExcelStringFormats.Bold10, blnWriteHeaders)
        End If

        'Store the cell formats.
        Cells.Clear()
        For Each iColumn As DataGridViewColumn In dg.Columns

            'returns false if DONOTPRINT is in the Tag property of a DataGridView Column.
            If CheckTag(iColumn, strTagFilter) = True Then
                If iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("N") Then
                    If iColumn.DefaultCellStyle.Format.Contains("0") Then
                        Cells.Add(New XMLExcelCell("", ExcelStringFormats.Decimal0, iColumn.DefaultCellStyle.Format))
                    Else
                        If iColumn.DefaultCellStyle.Format.Contains("1") Then
                            Cells.Add(New XMLExcelCell("", ExcelStringFormats.Decimal1, iColumn.DefaultCellStyle.Format))
                        Else
                            If iColumn.DefaultCellStyle.Format.Contains("2") Then
                                Cells.Add(New XMLExcelCell("", ExcelStringFormats.Decimal2, iColumn.DefaultCellStyle.Format))
                            Else
                                If iColumn.DefaultCellStyle.Format.Contains("3") Then
                                    Cells.Add(New XMLExcelCell("", ExcelStringFormats.Decimal3, iColumn.DefaultCellStyle.Format))
                                Else
                                    Cells.Add(New XMLExcelCell("", ExcelStringFormats.Decimal4, iColumn.DefaultCellStyle.Format))
                                End If
                            End If
                        End If
                    End If
                Else
                    If iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("C") Then
                        Cells.Add(New XMLExcelCell("", ExcelStringFormats.Currency2, iColumn.DefaultCellStyle.Format))
                    Else
                        If iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("P") Then
                            If iColumn.DefaultCellStyle.Format.Contains("0") Then
                                Cells.Add(New XMLExcelCell("", ExcelStringFormats.Percent0, iColumn.DefaultCellStyle.Format))
                            Else
                                If iColumn.DefaultCellStyle.Format.Contains("2") Then
                                    Cells.Add(New XMLExcelCell("", ExcelStringFormats.Percent2, iColumn.DefaultCellStyle.Format))
                                Else
                                    Cells.Add(New XMLExcelCell("", ExcelStringFormats.Percent4, iColumn.DefaultCellStyle.Format))
                                End If
                            End If
                        Else
                            Cells.Add(New XMLExcelCell("", ExcelStringFormats.NoStyle, iColumn.DefaultCellStyle.Format)) '  .Normal10))
                        End If
                    End If
                End If
            End If
        Next

        For Each TheRow As DataGridViewRow In dg.Rows

            'Do not process the last row of a r/w grid. It is a row of 0's.
            If TheRow.IsNewRow = False Then
                Dim iDataGridColumn As Integer
                iDataGridColumn = 0
                iCellsColumn = 0
                Dim iStyle As ExcelStringFormats = ExcelStringFormats.NormalInt10

                'Set the default row style. This is used instead of the cell style if this is NoStyle.
                Dim iRowStyle As ExcelStringFormats = ExcelStringFormats.Normal10

                '20120314 Reset the cell style of all the columns.

                For Each cell As XMLExcelCell In Cells
                    'cell.iStyle = ExcelStringFormats.Normal10
                    cell.blnBold = False
                Next

                'Set current row in datagridview bold.
                If Not dg.CurrentRow Is Nothing Then
                    If blnCurrentBold = True And TheRow.Index = dg.CurrentRow.Index Then

                        'If this is the currentrow and user wants bold then set the row style
                        'and choose NoStyle for the text cells. NoStyle ensures no cell style which
                        'then means that the row style will apply.
                        iRowStyle = ExcelStringFormats.BoldInt10
                        For Each cell As XMLExcelCell In Cells
                            If cell.iStyle = ExcelStringFormats.Normal10 Then
                                cell.iStyle = ExcelStringFormats.NoStyle
                            End If
                        Next
                    End If
                End If

                For Each iColumn As DataGridViewColumn In dg.Columns
                    Dim blnBold As Boolean = False

                    '20120314 Set the cell style to bold if the column is bold.
                    If Not iColumn.DefaultCellStyle.Font Is Nothing Then
                        If iColumn.DefaultCellStyle.Font.Bold Then
                            blnBold = True
                        End If
                    End If

                    If CheckTag(iColumn, strTagFilter) = True Then
                        If Not TheRow.Cells(iDataGridColumn).Value Is System.DBNull.Value Then

                            '20130522 TimeSpan was causing an exception.
                            Try
                                Cells(iCellsColumn).strValue = TheRow.Cells(iDataGridColumn).FormattedValue   '.Value
                            Catch ex As Exception
                                Dim ts As TimeSpan = TheRow.Cells(iDataGridColumn).Value
                                Cells(iCellsColumn).strValue = ts.ToString()
                            End Try
                            Cells(iCellsColumn).blnBold = blnBold
                        Else
                            If Cells(iCellsColumn).iStyle = iStyle Then
                                Cells(iCellsColumn).strValue = "0"
                            Else
                                Cells(iCellsColumn).strValue = ""
                            End If
                        End If
                        iCellsColumn = iCellsColumn + 1
                    End If
                    iDataGridColumn = iDataGridColumn + 1
                Next
                WriteStringToExcel(Cells, 1, False, iRowStyle, False)
            End If
        Next

        If blnWriteColumnWidths = True Then

            'Use the DataGridView Column width in Excel.
            Dim iColumn As Integer = 1
            For Each i As Integer In ColumnWidth
                If i > 12 Then
                    worksheet.Columns(iColumn).Width = (i - 12) / 7D + 1
                End If
                iColumn += 1
            Next
        End If

        ColumnWidth.Clear()
        For Each cell As XMLExcelCell In Cells
            cell.strValue = ""
            cell = Nothing
        Next

        Cells.Clear()
        ColumnWidth = Nothing
        Cells = Nothing

        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

End Class
