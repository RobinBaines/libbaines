'------------------------------------------------
'Name: Module ReadExcel.vb
'Function: 
'Copyright Baines 2007. All rights reserved.
'Notes:
'Modifications: Overloaded Open() including a format array, for UPLUS.
'20160609 Overloaded Open: Added an array of formats which is used only to identify dates at the moment. Put "Date" in the column to convert an Automation double date to
'a date in the current culture format.
'20241031 Converted to ClosedXML and removed all formatting.
'------------------------------------------------
Imports ClosedXML.Excel

Public Class ReadExcel
    Dim oBook As XLWorkbook
    Dim oSheet As IXLWorksheet

    Public Sub New()

    End Sub


    Public Function Open(strFileName As String) As Collection

        Dim ExcelRangeArray As Collection
        ExcelRangeArray = New Collection
        Using oBook = New XLWorkbook(strFileName)
            oSheet = oBook.Worksheet(1)
            Dim Column As Integer = 1
            Dim Row As Integer = 1
            Dim rowCount As Integer = oSheet.LastRowUsed().RowNumber()
            Dim columnCount As Integer = oSheet.LastColumnUsed().ColumnNumber()

            'oSheet.Rows.ClearFormats()
            While (Row <= rowCount)
                Dim strArray(columnCount) As String
                While (Column <= columnCount)
                    Dim obj As Object
                    obj = oSheet.Cell(Row, Column).GetString()
                    strArray(Column - 1) = obj
                    Column += 1
                End While
                ExcelRangeArray.Add(strArray)
                Row += 1
                Column = 1
            End While
        End Using
        Return ExcelRangeArray
    End Function
End Class
