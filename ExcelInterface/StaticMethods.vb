'------------------------------------------------
'Name: Module StaticMethods.vb
'Function: 
'Copyright Baines 2007. All rights reserved.
'Notes:
'Modifications: 
'20110701 RPB added CloseExcelBook which is a standalone function which converts an Excel 2003 XML file to a workbook xls file.
'------------------------------------------------
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data.SqlTypes
Imports Microsoft.VisualBasic.FileIO.FileSystem
Imports Microsoft.Vbe.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Xml

Namespace ExcelInterfaceStatic
    Public Module StaticMethods
        Dim wb As Workbook
        Dim doc As New XmlDocument
        Dim nsmgr As XmlNamespaceManager

        Public Delegate Sub UpdateUICB(ByVal iProgress As Integer, ByVal strDescription As String, ByVal blnEndOfProcess As Boolean)
        Dim ShowProgress As UpdateUICB
        Dim theForm As Form

        'Called when parent binding source has processed a new row.
        Public Delegate Sub NewRecordUICB(ByVal bsParent As BindingSource)
        Dim NewRecord As NewRecordUICB

        '----------------------------------------------------------------
        'CreateWorkBookFromTemplate
        'Entry point to create excel book with a sheet per parent record. 
        '----------------------------------------------------------------
        Public Function CreateWorkBookFromTemplate(ByVal _theForm As Form, ByVal _ShowProgress As UpdateUICB, _
        ByVal _NewRecord As NewRecordUICB, _
        ByVal strTemplate As String, _
        ByVal strFileName As String, ByVal strBackGroundFileName As String, _
        ByVal tParent As System.Data.DataTable, _
        ByVal tChild As System.Data.DataTable, ByVal bsParent As BindingSource, _
        ByVal bsChild As BindingSource, _
        ByVal dgChild As DataGridView, _
        ByVal iSheetnameIndex As Integer, _
        ByVal blnPDF As Boolean, ByVal blnRunMacro As Boolean, _
        ByVal blnShowExcel As Boolean, _
        ByVal blnDeleteXML As Boolean)
            Dim strFileSaved As String = ""
            Try
                ShowProgress = _ShowProgress
                NewRecord = _NewRecord
                theForm = _theForm
                Dim root As XmlElement = LoadTemplate(strTemplate)
                Dim nSheet As XmlNode
                nSheet = root.SelectSingleNode("s:Worksheet", nsmgr) 'ss:Name='Blad1'")
                If Not nSheet Is Nothing Then
                    CopyWorksheet(tParent, tChild, bsParent, bsChild, dgChild, iSheetnameIndex, root, nSheet)
                End If
                strFileSaved = CloseDoc(strFileName, strBackGroundFileName, blnPDF, blnRunMacro, blnShowExcel, blnDeleteXML)
            Catch ex As Exception
            End Try
            Return strFileSaved
        End Function

        '----------------------------------------------------------------
        'CreateWorkBookSingleSheetFromTemplate
        'Entry point to create excel book with a sheet per parent record. Is used in SFC.
        '----------------------------------------------------------------
        Public Function CreateWorkBookSingleSheetFromTemplate(ByVal _theForm As Form, ByVal _ShowProgress As UpdateUICB, _
        ByVal strTemplate As String, _
        ByVal strFileName As String, ByVal strBackGroundFileName As String, _
        ByVal tParent As System.Data.DataTable, _
        ByVal tChild As System.Data.DataTable, ByVal bsParent As BindingSource, _
        ByVal bsChild As BindingSource, _
        ByVal dgChild As DataGridView, _
        ByVal iSheetnameIndex As Integer, _
        ByVal blnPDF As Boolean, ByVal blnRunMacro As Boolean, _
        ByVal blnShowExcel As Boolean, _
        ByVal blnDeleteXML As Boolean)
            Dim strFileSaved As String = ""
            Try
                ShowProgress = _ShowProgress
                theForm = _theForm
                Dim root As XmlElement = LoadTemplate(strTemplate)
                Dim nSheet As XmlNode
                nSheet = root.SelectSingleNode("s:Worksheet", nsmgr)
                If Not nSheet Is Nothing Then
                    CopyAWorksheet(tParent, tChild, bsParent, bsChild, dgChild, iSheetnameIndex, root, nSheet, 0)
                End If
                strFileSaved = CloseDoc(strFileName, strBackGroundFileName, blnPDF, blnRunMacro, blnShowExcel, blnDeleteXML)
            Catch ex As Exception
            End Try
            Return strFileSaved
        End Function

        Private Function LoadTemplate(ByVal strTemplate As String) As XmlElement
            doc.Load(strTemplate)
            nsmgr = New XmlNamespaceManager(doc.NameTable)
            nsmgr.AddNamespace("s", "urn:schemas-microsoft-com:office:spreadsheet")
            nsmgr.AddNamespace("", "urn:schemas-microsoft-com:office:spreadsheet")
            Dim root As XmlElement = doc.DocumentElement
            Dim nSheet As XmlNode
            nSheet = root.SelectSingleNode("s:Worksheet", nsmgr)
            Return root
        End Function

        Private Sub InsertHeaderData(ByVal t As System.Data.DataTable, ByVal sd As DataRowView, ByVal nCells As XmlNodeList)
            Dim dr As DataRow = sd.Row
            For Each nCell As XmlNode In nCells
                If Not nCell Is Nothing Then
                    For i As Integer = 0 To t.Columns.Count - 1
                        Try
                            'Debug.Print(nCell.InnerText & " " & t.Columns(i).ColumnName)
                            If t.Columns(i).ColumnName.ToLower = nCell.InnerText.ToLower Then
                                Dim dNodes As XmlNodeList = nCell.SelectNodes("s:Data", nsmgr)
                                For Each dN As XmlNode In dNodes
                                    Try
                                        'error if DBNull
                                        dN.InnerText = sd.Row.Item(i)
                                    Catch ex As Exception

                                    End Try

                                Next
                                Exit For
                            End If
                        Catch ex As Exception
                        End Try
                    Next
                End If
            Next
        End Sub
        '
        Private Function strGetColumnFormat(ByVal strColumnName As String) As String
            Return "String"
        End Function

        'Get the header cells. If the header cell contains a field name enter the value in the target cell: nCell.
        Private Function InsertDetailData(ByVal t As System.Data.DataTable, ByVal bs As BindingSource, ByVal nHeaderRow As XmlNode, ByVal nCells As XmlNodeList) As Boolean
            Dim nHeaders As XmlNodeList = nHeaderRow.SelectNodes("s:Cell", nsmgr)
            Dim sdChild As DataRowView = bs.Current
            Dim iCellCount As Integer = 0
            For Each nHeaderCell As XmlNode In nHeaders
                If Not nHeaderCell Is Nothing Then
                    Dim nTargetCell As XmlNode = nCells.Item(iCellCount)

                    'Then find the same column in the data table.
                    For i As Integer = 0 To t.Columns.Count - 1
                        If t.Columns(i).ColumnName.ToLower = nHeaderCell.InnerText.ToLower Then

                            'See whether there is data in the 'empty' detail row.
                            'If not create the element. If there is (there shouldn't be) just update it.
                            Dim dNodes As XmlNodeList = nTargetCell.SelectNodes("s:Data", nsmgr)
                            If dNodes.Count = 0 Then
                                Dim elem As XmlElement = doc.CreateElement("", "Data", "urn:schemas-microsoft-com:office:spreadsheet")
                                Dim aAt As XmlAttribute = doc.CreateAttribute("ss", "Type", "urn:schemas-microsoft-com:office:spreadsheet")
                                aAt.Value = strGetColumnFormat(t.Columns(i).ColumnName)
                                elem.SetAttributeNode(aAt)
                                elem.InnerText = sdChild.Row.Item(i)
                                nTargetCell.AppendChild(elem)
                            Else
                                For Each dN As XmlNode In dNodes
                                    'Don't change the format.
                                    Try
                                        dN.InnerText = sdChild.Row.Item(i)
                                    Catch ex As Exception

                                    End Try

                                Next
                            End If
                            Exit For
                        End If
                    Next
                End If
                iCellCount = iCellCount + 1
            Next
        End Function

        'Insert and fill rows for each child record.
        'Clone the row from nRow.
        Private Sub InsertDetailRows(ByVal t As System.Data.DataTable, ByVal bs As BindingSource, ByVal nTable As XmlNode, ByVal nHeaderRow As XmlNode, ByVal nRow As XmlNode)
            Dim nCells As XmlNodeList

            'Make a copy of the row to use as fresh clone.
            Dim nRowClone As XmlNode = nRow.Clone
            Dim nLastRow As XmlNode = nRow
            nCells = nRow.SelectNodes("s:Cell", nsmgr)
            If bs.Count > 0 Then
                InsertDetailData(t, bs, nHeaderRow, nCells)
                For i As Integer = 1 To bs.Count - 1
                    Dim iPosition As Integer = bs.Position
                    bs.MoveNext()

                    'Not needed but added for safety.
                    If iPosition = bs.Position Then
                        Exit For
                    End If

                    Dim nRowNewClone As XmlNode = nRowClone.Clone
                    nCells = nRowNewClone.SelectNodes("s:Cell", nsmgr)
                    InsertDetailData(t, bs, nHeaderRow, nCells)
                    nTable.InsertAfter(nRowNewClone, nLastRow)

                    'Save this row as reference for the next InsertAfter.
                    nLastRow = nRowNewClone
                Next
            Else

            End If
        End Sub

        'Use the header of the datagrid to update the field names in the order_detail header row.
        Private Sub ShowCorrectHeaders(ByVal dg As DataGridView, ByVal nHeaderRow As XmlNode)

            Dim nHeaders As XmlNodeList = nHeaderRow.SelectNodes("s:Cell", nsmgr)
            Dim iCellCount As Integer = 0
            For Each nCell As XmlNode In nHeaders

                'Then find the same column in the data table.
                For i As Integer = 0 To dg.Columns.Count - 1
                    If dg.Columns(i).Name.ToLower = nCell.InnerText.ToLower Then
                        Dim dNodes As XmlNodeList = nCell.SelectNodes("s:Data", nsmgr)
                        For Each dN As XmlNode In dNodes
                            'Debug.Print(dN.InnerText)
                            dN.InnerText = dg.Columns(i).HeaderText
                            'Debug.Print(dN.InnerText)
                        Next
                    End If
                Next
            Next
        End Sub

        'Header fields are filled where they are found.
        'The named field only has to start with the header text. This allows the header functionality to be switched on and off.
        Const HEADER = "header"

        'Ignore this text and leave it as it is.
        'Named field only has to start with this.
        Const NORMAL_TXT = "normal_txt"

        'Insert the datagrid. This named field should/can only occur once.
        Const DETAIL_TABLE = "detail_table"
        Const FILL_DETAIL = "order_fill_detail"

        Private Sub FillWorkSheet(ByVal tParent As System.Data.DataTable, _
        ByVal tChild As System.Data.DataTable, ByVal bs As BindingSource, _
        ByVal dg As DataGridView, ByVal newSheet As XmlNode, ByVal sd As DataRowView, ByVal strSheetName As String)

            Try

                'Set the name of the sheet.
                newSheet.Attributes("ss:Name").InnerText = strSheetName
                Dim strStatus As String = ""

                'Get the table corresponding to the original sheet.
                Dim nTable As XmlNode = newSheet.SelectSingleNode("s:Table", nsmgr)
                If Not nTable Is Nothing Then
                    Dim nlRows As XmlNodeList = nTable.SelectNodes("s:Row", nsmgr)
                    Dim nHeaderRow As XmlNode = Nothing

                    'Iterate through the rows of the table.
                    For Each nRow As XmlNode In nlRows
                        Dim nCells As XmlNodeList
                        nCells = nRow.SelectNodes("s:Cell", nsmgr)
                        For Each nCell As XmlNode In nCells
                            If Not nCell Is Nothing Then

                                'Get a named cell if there is one (see Excel Formulas Define names).
                                Dim nNamedCell As XmlNode = nCell.SelectSingleNode("s:NamedCell", nsmgr)
                                If Not nNamedCell Is Nothing Then
                                    Dim xAts As XmlAttributeCollection = nNamedCell.Attributes
                                    For Each nAt As XmlAttribute In xAts
                                        If nAt.InnerText().Length >= HEADER.Length Then
                                            If nAt.InnerText().Substring(0, HEADER.Length) = HEADER Then
                                                strStatus = HEADER  'nAt.InnerText().Substring(0, HEADER.Length)
                                                ' MsgBox("found order_header")
                                            End If
                                        End If
                                        If nAt.InnerText().Length >= NORMAL_TXT.Length Then
                                            If nAt.InnerText().Substring(0, NORMAL_TXT.Length) = NORMAL_TXT Then
                                                strStatus = NORMAL_TXT  'nAt.InnerText().Substring(0, NORMAL_TXT.Length)
                                                ' MsgBox("found order_header")
                                            End If
                                        End If
                                        If nAt.InnerText().Length >= DETAIL_TABLE.Length Then
                                            If nAt.InnerText() = DETAIL_TABLE Then
                                                nHeaderRow = nRow   '.Clone()
                                                If bs.Count > 0 Then bs.MoveFirst()
                                                strStatus = DETAIL_TABLE 'nAt.InnerText()
                                                'MsgBox("found order_detail")
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        Next

                        'The DETAIL_TABLE status is set when the header of the table is found.
                        'After filling in the header fields the status is set to the following to allow the table rows to be filled.
                        If strStatus = FILL_DETAIL Then
                            'comes here with the detail row.
                            InsertDetailRows(tChild, bs, nTable, nHeaderRow, nRow)
                            strStatus = ""
                        End If
                        If strStatus = HEADER Then
                            'stay in this status until it is set to order_status.
                            InsertHeaderData(tParent, sd, nCells)
                        Else
                            If strStatus = DETAIL_TABLE Then
                                strStatus = FILL_DETAIL
                            End If
                        End If
                    Next

                    'Having created the sheet go back to the detail header row and show the translated headers.
                    If Not nHeaderRow Is Nothing Then
                        ShowCorrectHeaders(dg, nHeaderRow)
                    End If

                    'Change this attribute to allow for added rows. Fails if not done. 
                    Dim strRows = nTable.Attributes.GetNamedItem("ss:ExpandedRowCount").Value
                    Try
                        Dim iRows As Integer
                        If bs.Count > 0 Then
                            iRows = System.Convert.ToInt32(strRows) + bs.Count - 1
                        Else
                            iRows = System.Convert.ToInt32(strRows)
                        End If

                        nTable.Attributes.GetNamedItem("ss:ExpandedRowCount").Value = iRows.ToString()
                    Catch ex As Exception
                        nTable.Attributes.RemoveNamedItem("ss:ExpandedRowCount")
                    End Try
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Sub

        'Copy original sheet for each order.
        'The fill the sheet.
        Private Function CopyAWorksheet(ByVal tParent As System.Data.DataTable, _
        ByVal tChild As System.Data.DataTable, _
        ByVal bsParent As BindingSource, _
        ByVal bsChild As BindingSource, _
        ByVal dgChild As DataGridView, _
        ByVal iSheetnameIndex As Integer, _
        ByVal root As XmlElement, _
        ByVal nSheet As XmlNode, _
        ByVal iPosition As Integer) As String
            Dim newSheet As XmlNode
            If iPosition = 0 Then
                newSheet = nSheet
            Else
                Dim CopySheet As XmlNode = nSheet.Clone()
                newSheet = CopySheet.Clone()
            End If
            Dim sd As DataRowView = bsParent.Current
            Dim strSheetName = sd.Row.Item(iSheetnameIndex)
            FillWorkSheet(tParent, tChild, bsChild, dgChild, newSheet, sd, strSheetName)
            If iPosition <> 0 Then
                root.AppendChild(newSheet)
            End If
            Return strSheetName
        End Function

        'Copy original sheet for each order.
        'The fill the sheet.
        Private Sub CopyWorksheet(ByVal tParent As System.Data.DataTable, _
        ByVal tChild As System.Data.DataTable, ByVal bsParent As BindingSource, _
        ByVal bsChild As BindingSource, _
        ByVal dgChild As DataGridView, _
        ByVal iSheetnameIndex As Integer, ByVal root As XmlElement, ByVal nSheet As XmlNode)
            Dim iCount = bsParent.Count
            bsParent.MoveFirst()
            Dim iPosition As Integer = -1
            Dim iFirstPosition As Integer = bsParent.Position
            Do While iPosition <> bsParent.Position
                iPosition = bsParent.Position
                Dim strSheetName = CopyAWorksheet(tParent, tChild, bsParent, bsChild, dgChild, iSheetnameIndex, root, nSheet, iPosition)
                ShowProgress(100 * iPosition / iCount, "Sheet = " & strSheetName, False)
                NewRecord(bsParent)
                bsParent.MoveNext()
            Loop
        End Sub

        Private Function CloseDoc(ByVal strFileName As String, ByVal strBackGroundFileName As String, _
        ByVal blnPDF As Boolean, _
        ByVal blnRunMacro As Boolean, _
        ByVal blnShowExcel As Boolean, _
        ByVal blnDeleteXML As Boolean) As String
            Dim strFileSaved As String = ""
            Try
                Dim strPathFileName = MakeUniquePath(0, True, ExcelInterface.Paths.Local, "", strFileName, ".xml")
                'MsgBox("strPathFileName = " & strPathFileName)
                doc.Save(strPathFileName)
                strFileSaved = CloseExcelBook(strFileName, strPathFileName, blnPDF, blnShowExcel, blnDeleteXML, strBackGroundFileName, blnRunMacro)
                nsmgr = Nothing
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Return strFileSaved
        End Function

        '20110701 CloseExcelBook is a standalone function which converts an Excel 2003 XML file to a workbook xls file.
        Public Function CloseExcelBook(ByVal strFileName As String, ByVal strPathFileName As String, ByVal blnPDF As Boolean, _
            ByVal blnShowExcel As Boolean, _
            ByVal blnDeleteXML As Boolean, ByVal BackGroundFileName As String, _
        ByVal blnRunMacro As Boolean) As String
            Dim strFileSaved As String = ""
            Try
                'Open the XML file in Excel.
                If My.Computer.FileSystem.FileExists(strPathFileName) Then
                    'MsgBox("File exists = " & strPathFileName)
                    ShowProgress(100, "Loading XLS", False)
                    Dim oExcel As New Microsoft.Office.Interop.Excel.Application
                    oExcel.Visible = True
                    oExcel.UserControl = True
                    Dim oldCI As System.Globalization.CultureInfo = _
                        System.Threading.Thread.CurrentThread.CurrentCulture

                    System.Threading.Thread.CurrentThread.CurrentCulture = _
                                    New System.Globalization.CultureInfo("en-US")

                    Dim oBooks As Workbooks
                    oBooks = oExcel.Workbooks
                    wb = oBooks.Open(strPathFileName)
                    If Not wb Is Nothing Then
                        If BackGroundFileName.Length > 0 Then
                            ShowProgress(100, "Adding Background", False)
                            For Each sh As Worksheet In wb.Sheets
                                sh.PageSetup.CenterHeaderPicture.Filename = BackGroundFileName
                                'With sh.PageSetup
                                '    .LeftHeader = ""
                                '    .CenterHeader = "&G"
                                '    .RightHeader = ""
                                '    .LeftFooter = ""
                                '    .CenterFooter = ""
                                '    .RightFooter = ""
                                '    .LeftMargin = oExcel.InchesToPoints(0.7)
                                '    .RightMargin = oExcel.InchesToPoints(0.7)
                                '    .TopMargin = oExcel.InchesToPoints(0.75)
                                '    .BottomMargin = oExcel.InchesToPoints(0.75)
                                '    .HeaderMargin = oExcel.InchesToPoints(0.3)
                                '    .FooterMargin = oExcel.InchesToPoints(0.3)
                                '    .Zoom = 82
                                '    ' .PrintErrors = oExcel.xlPrintErrorsDisplayed
                                '    .OddAndEvenPagesHeaderFooter = False
                                '    .DifferentFirstPageHeaderFooter = False
                                '    .ScaleWithDocHeaderFooter = True
                                '    .AlignMarginsHeaderFooter = True
                                '    .EvenPage.LeftHeader.Text = ""
                                '    .EvenPage.CenterHeader.Text = ""
                                '    .EvenPage.RightHeader.Text = ""
                                '    .EvenPage.LeftFooter.Text = ""
                                '    .EvenPage.CenterFooter.Text = ""
                                '    .EvenPage.RightFooter.Text = ""
                                '    .FirstPage.LeftHeader.Text = ""
                                '    .FirstPage.CenterHeader.Text = ""
                                '    .FirstPage.RightHeader.Text = ""
                                '    .FirstPage.LeftFooter.Text = ""
                                '    .FirstPage.CenterFooter.Text = ""
                                '    .FirstPage.RightFooter.Text = ""
                                'End With
                            Next
                        End If

                        'Autofit is not defined in the xml or at least does not trigger when data is added.
                        'Do it here. The macro defined below was also set up to do it but is no longer necessary so function is called with blnRunMacro = false.
                        For Each oSheet As Worksheet In wb.Sheets
                            oSheet.Rows.AutoFit()
                        Next

                        If blnRunMacro = True Then
                            Dim oModule As VBComponent

                            'This will fail if the user has not activated 'trust access to the VBA project object model'.
                            Try
                                ShowProgress(100, "Running macro", False)
                                oModule = wb.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule)
                                Dim sCode As String
                                sCode = "sub VBAMacro()" & vbCr & _
                                "Cells.Select" & vbCr & _
                                "Selection.Rows.AutoFit" & vbCr & _
                                "Range(""A2"").Select" & vbCr & _
                                   "end sub"
                                ''"   msgbox ""VBA Macro called"" " & vbCr & _
                                '' "end sub"

                                ' Add the VBA macro to the new code module.
                                oModule.CodeModule.AddFromString(sCode)
                                oExcel.Run("VBAMacro")
                                wb.VBProject.VBComponents.Remove(oModule)
                                'oModule.CodeModule.DeleteLines(0)
                                oModule = Nothing
                            Catch ex As Exception

                            End Try
                        End If

                        'Just need the file name without the path.

                        If blnPDF = False Then
                            strFileSaved = MakeUniquePath(0, True, ExcelInterface.Paths.Local, "", strFileName, ".xlsx")
                            strFileSaved = strFileSaved.Substring(0, strFileSaved.Length() - Len(".xlsx"))
                            ShowProgress(100, "Saving Excel file " & strFileSaved & ".xlsx", False)

                            'The saveas converts the xml to an xlsx file.
                            wb.SaveAs(strFileSaved, 51)
                        Else
                            strFileSaved = MakeUniquePath(0, True, ExcelInterface.Paths.Local, "", strFileName, ".pdf")
                            'strExcelFilename = strExcelFilename.Substring(0, strExcelFilename.Length() - Len(".pdf"))
                            ShowProgress(100, "Saving pdf file " & strFileSaved, False)
                            'Dim strPDFFilename = strExcelFilename & ".pdf"
                            Dim paramExportFilePath As String = strFileSaved
                            Dim paramExportFormat As XlFixedFormatType = XlFixedFormatType.xlTypePDF
                            Dim paramExportQuality As XlFixedFormatQuality = XlFixedFormatQuality.xlQualityStandard
                            Dim paramOpenAfterPublish As Boolean = False
                            Dim paramIncludeDocProps As Boolean = True
                            Dim paramIgnorePrintAreas As Boolean = True
                            Dim paramFromPage As Object = Type.Missing
                            Dim paramToPage As Object = Type.Missing
                            wb.ExportAsFixedFormat(paramExportFormat, _
                                        paramExportFilePath, paramExportQuality, _
                                        paramIncludeDocProps, paramIgnorePrintAreas, _
                                        paramFromPage, paramToPage, paramOpenAfterPublish)
                            ShowProgress(100, "Saving file " & strFileSaved, False)
                            Dim strFileRootName = strFileSaved.Substring(0, strFileSaved.Length() - Len(".pdf"))
                            wb.SaveAs(strFileRootName, 51)
                        End If
                        ShowProgress(100, "Cleaning up", False)
                        If blnShowExcel = True Then
                            wb.Application.UserControl = True
                            If Not wb Is Nothing Then
                                wb = Nothing
                            End If
                        Else
                            If Not wb Is Nothing Then
                                wb.Close(False)
                                wb = Nothing
                            End If

                            ' Quit Excel and release the ApplicationClass object.
                            If Not oExcel Is Nothing Then
                                oExcel.Quit()
                                oExcel = Nothing
                            End If
                            GC.Collect()
                            GC.WaitForPendingFinalizers()
                            GC.Collect()
                            GC.WaitForPendingFinalizers()
                        End If
                        If blnDeleteXML = True Then
                            My.Computer.FileSystem.DeleteFile(strPathFileName)
                        End If
                    End If
                    System.Threading.Thread.CurrentThread.CurrentCulture = oldCI
                Else
                    MsgBox("File does not exist = " & strPathFileName)
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Return strFileSaved
        End Function

        Private Function GetFilePath(ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal strFileName As String, ByVal strExtension As String) As String
            Return GetDirectory(pPath, strSubDirectory) & strFileName & strExtension
        End Function

        Dim strLocalOutput = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        Dim strNetworkOutput = My.Computer.FileSystem.SpecialDirectories.MyDocuments

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

        Public Function MakeUniquePath(ByVal iSeed As Integer, ByVal blnUnique As Boolean, ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal strFileName As String, ByVal strExtension As String) As String

            'Recursive call which increments the filename until the filename does not exist.
            Dim strFilePath As String
            If iSeed = 0 Then
                strFilePath = GetFilePath(pPath, strSubDirectory, strFileName, strExtension)
            Else
                strFilePath = GetFilePath(pPath, strSubDirectory, strFileName & "_" & iSeed.ToString("0000"), strExtension)
            End If
            If FileExists(strFilePath) = True Then
                If blnUnique = False Or iSeed > 100 Then
                    My.Computer.FileSystem.DeleteFile(strFilePath)
                Else
                    strFilePath = MakeUniquePath(iSeed + 1, True, pPath, strSubDirectory, strFileName, strExtension)
                End If
            End If
            Return strFilePath
        End Function
    End Module
End Namespace
