'------------------------------------------------
'Name: Module clXMLExcelCell.vb.
'Function: 
'Copyright Robin Baines 2024. All rights reserved.
'Created Oct 2024.
'Notes: 
'Modifications: 
'------------------------------------------------
Public Class XMLExcelCell
    Public strValue As String
    Public iStyle As ExcelStringFormats
    Public blnBold As Boolean
    Public NumberFormat As String
    Public Sub New(ByVal _strValue As String, ByVal _iStyle As ExcelStringFormats, ByVal _NumberFormat As String)
        strValue = _strValue
        iStyle = _iStyle
        blnBold = False
        NumberFormat = _NumberFormat
    End Sub
    Public Sub New(ByVal _strValue As String, ByVal _iStyle As ExcelStringFormats)
        strValue = _strValue
        iStyle = _iStyle
        blnBold = False
        NumberFormat = ""
    End Sub
End Class
