'------------------------------------------------
'Name: Module InterfaceToExcel.vb.
'Function: 
'Copyright Robin Baines 2007. All rights reserved.
'Created Dec 2007.
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Windows.Forms
Public Interface InterfaceToExcel
    Sub OpenExcelBook(ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal strFileName As String, _
            ByVal blnUnique As Boolean)
    Sub OpenExcelBook(ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal strFileName As String, _
            ByVal blnUnique As Boolean, ByVal blnVisible As Boolean)
    Sub OpenExcelBook(ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal strFileName As String, _
            ByVal strHeaderText As String, ByVal blnLandscape As Boolean)
    Function SetActivePrinter(ByVal strPrinter As String) As Boolean
    Sub NewSheet(ByVal strHeaderText As String, ByVal blnLandScape As Boolean, ByVal blnShowHeader As Boolean, _
        ByVal dLeftMargin As Double, ByVal dRightMargin As Double, ByVal dTopMargin As Double, ByVal dBottomMargin As Double, _
        ByVal dHeaderMargin As Double, _
        ByVal dFooterMargin As Double, ByVal iPaperSize As Double, ByVal blnFitToPage As Boolean, _
        ByVal blnRepeatTopRow As Boolean, ByVal iPrintQuality As Integer, ByVal iZoom As Integer)
    Sub NewSheet(ByVal strHeaderText As String, ByVal blnLandScape As Boolean, ByVal blnShowHeader As Boolean)
    Sub CloseExcelBook()
    Sub CloseExcelBookAndExitExcel()
    Sub SetBorders(ByVal iBorder As BorderWeight, ByVal iFirstRow As Integer, _
        ByVal iLastRow As Integer, _
        ByVal iFirstColumn As Integer, _
        ByVal iLastColumn As Integer)
    Sub MakeASum(ByVal iAtRow As Integer, ByVal iAtColumn As Integer, _
            ByVal iFromRow As Integer, ByVal iToRow As Integer, ByVal dDivideBy As Double)
    Sub CreateFormula(ByVal iAtRow As Integer, ByVal iAtColumn As Integer, _
    ByVal strFormula As String)
    Sub SetColumnFormat(ByVal strFormat As String, ByVal iFirstRow As Integer, _
        ByVal iFirstColumn As Integer, _
        ByVal iLastColumn As Integer)
    Sub SetColumnFormat(ByVal strFormat As String, ByVal iFirstRow As Integer, _
        ByVal iLastRow As Integer, _
        ByVal iFirstColumn As Integer, _
        ByVal iLastColumn As Integer)
    Sub SetColumnBold(ByVal blnBold As Boolean, ByVal iFirstRow As Integer, ByVal iColumn As Integer)
    Sub SetColumnWidth(ByVal iColumn As Integer, ByVal iWidth As Integer)
    Sub SetCellFont(ByVal fFormat As ExcelStringFormats, ByVal iRow As Integer, ByVal iCol As Integer)
    Sub WriteDataGrid(ByVal dg As DataGridView, ByVal strTagFilter As String, ByVal blnCurrentBold As Boolean, _
        ByVal iFirstColumn As Integer, ByVal blnAutofit As Boolean)
    Sub AdjustHeightOfRow(ByVal dHeight As Double)
    Sub AdjustHeightOfRow(ByVal iColumn As Integer, ByVal dHeight As Double)
    Sub AdjustWidthOfColumn(ByVal iColumn As Integer, ByVal dWidth As Double)
    Sub SetRow(ByVal iRow As Integer)
    Sub SetVerticalPageBreak(ByVal iColumn As Integer)
    Sub SetHorizontalPageBreak(ByVal iRow As Integer)
    Sub WriteStringToExcelAndMerge(ByVal strT As String, ByVal fFormat As ExcelStringFormats, ByVal iNrColumns As Integer)
    Sub WriteStringToExcel(ByVal strT As String, ByVal fFormat As ExcelStringFormats)
    Sub WriteStringToExcel(ByVal strT As String, ByVal fFormat As ExcelStringFormats, _
        ByVal iBorder As BorderWeight)
    Sub SetRowColour(ByVal iIndex As Integer)
    Sub SetCellColour(ByVal iColumn As Integer, ByVal iIndex As Integer)
    Function WriteStringToExcel(ByVal strT As String, _
        ByVal fFormat As ExcelStringFormats, _
        ByVal iStartColumn As Integer, _
        ByVal blnRightAlign As Boolean, _
        ByVal iBorder As BorderWeight) As Integer
    Sub strCreateExcelSheet(ByVal dg As DataGridView, ByVal strFileName As String, ByVal strHeaderText As String, _
    ByVal strTagFilter As String)
    Sub SetPrintArea()
    Sub SetPrintArea(ByVal iFirstRow As Integer, ByVal iFirstColumn As Integer, _
        ByVal iLastRow As Integer, ByVal iLastColumn As Integer)
    Sub SetRowFont(ByVal rCount As Long, ByVal fFormat As ExcelStringFormats)
    Sub PrintMasterChild(ByVal dgMaster As DataGridView, ByVal dgChild As DataGridView, _
        ByVal strSubDirectory As String, ByVal strFileName As String, ByVal strHeader As String, ByVal strTagFilter As String)
    Sub InsertPicture(ByVal strImageName As String)
    Sub SetBorder()
    Sub SetAutofit()
    Sub SetFitTo1Page()
    Function strPrecision(ByVal strFormat As String) As String

End Interface
