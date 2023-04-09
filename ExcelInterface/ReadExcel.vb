'------------------------------------------------
'Name: Module ReadExcel.vb
'Function: 
'Copyright Baines 2007. All rights reserved.
'Notes:
'Modifications: Overloaded Open() including a format array, for UPLUS.
'20160609 Overloaded Open: Added an array of formats which is used only to identify dates at the moment. Put "Date" in the column to convert an Automation double date to
'a date in the current culture format.
'------------------------------------------------
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data.SqlTypes
Imports Microsoft.VisualBasic.FileIO.FileSystem
Imports Microsoft.Office.Interop
Imports System.Collections
Public Class ReadExcel
    Dim xlApp As Excel.Application
    Dim oBook As Excel.Workbook
    Dim oSheet As Excel.Worksheet
    'Dim oBooks As Excel.Workbooks

    Public Sub New()

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

    '20160609 Added an array of formats which is used only to identify dates at the moment. Put "Date" in the column to convert an Automation double date to
    'a date in the current culture format.
    Public Function Open(strFileName As String, LastColumnLetter As String) As Object
        '  Public Function Open(strFileName As String, iMaxLines As Integer, iMaxColumns As Integer, LastColumnLetter As String) As Collection
        Dim A As Collection
        A = New Collection
        xlApp = New Excel.Application
        oBook = xlApp.Workbooks.Open(
        strFileName, 0, True, 5,
          "", "", True, Excel.XlPlatform.xlWindows, "\t", False, False,
          0, True)

        oSheet = oBook.Worksheets(1)
        'oSheet.Columns.ClearFormats()
        oSheet.Rows.ClearFormats()
        'Dim last As Excel.Range = oSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing)
        Dim ExcelRangeArray(0, 0) As Object
        Dim range As Excel.Range = oSheet.Range("A1", LastColumnLetter + oSheet.UsedRange.Rows.Count.ToString())
        ExcelRangeArray = range.Value
        'A = range.Value

        xlApp.DisplayAlerts = False
        xlApp.ActiveWorkbook.Close()
        xlApp.Quit()
        NAR(oSheet)
        NAR(oBook)

        'Not sure about this but by not calling the Yes/No/Cancel dialog is avoided.
        'xlApp.Quit()
        NAR(xlApp)
        oSheet = Nothing
        oBook = Nothing
        xlApp = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Return ExcelRangeArray
    End Function



    Public Function Open(strFileName As String, iMaxLines As Integer, iMaxColumns As Integer, LastColumnLetter As String) As Collection
        Dim A As Collection
        A = New Collection
        xlApp = New Excel.Application
        oBook = xlApp.Workbooks.Open(
        strFileName, 0, True, 5,
          "", "", True, Excel.XlPlatform.xlWindows, "\t", False, False,
          0, True)

        oSheet = oBook.Worksheets(1)
        oSheet.Columns.ClearFormats()
        oSheet.Rows.ClearFormats()
        Dim last As Excel.Range = oSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing)
        Dim i As Integer = 1
        While i < iMaxLines
            Dim range As Excel.Range = oSheet.Range("A" + i.ToString(), LastColumnLetter + i.ToString())
            Dim myvalues As System.Array = range.Cells.Value
            Dim obj As String
            Dim y As Integer = 0
            Dim strArray(iMaxColumns) As String
            For Each obj In myvalues
                strArray(y) = obj
                If y = 0 Then
                    If obj Is Nothing Then
                        Exit While
                    End If
                    If obj.Length = 0 Then
                        Exit While
                    End If
                End If

                y += 1
                If y = iMaxColumns Then Exit For
            Next
            A.Add(strArray)
            i += 1
            If i = iMaxLines Then Exit While
        End While
        xlApp.DisplayAlerts = False
        xlApp.ActiveWorkbook.Close()
        xlApp.Quit()

        NAR(oSheet)
        NAR(oBook)


        'Not sure about this but by not calling the Yes/No/Cancel dialog is avoided.
        'xlApp.Quit()
        NAR(xlApp)
        oSheet = Nothing
        oBook = Nothing
        xlApp = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Return A
    End Function


    '20160609 Overloaded Open: Added an array of formats which is used only to identify dates at the moment. Put "Date" in the column to convert an Automation double date to
    'a date in the current culture format.
    Public Function Open(strFileName As String, iMaxLines As Integer, strFormats() As String, iMaxColumns As Integer, LastColumnLetter As String) As Collection
        '  Public Function Open(strFileName As String, iMaxLines As Integer, iMaxColumns As Integer, LastColumnLetter As String) As Collection
        Dim A As Collection
        A = New Collection
        xlApp = New Excel.Application
        oBook = xlApp.Workbooks.Open(
        strFileName, 0, True, 5,
          "", "", True, Excel.XlPlatform.xlWindows, "\t", False, False,
          0, True)

        oSheet = oBook.Worksheets(1)
        'oSheet.Columns.ClearFormats()
        oSheet.Rows.ClearFormats()
        'Dim last As Excel.Range = oSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing)
        Dim i As Integer = 1

        While i < iMaxLines
            Dim range As Excel.Range = oSheet.Range("A" + i.ToString(), LastColumnLetter + i.ToString())
            Dim myvalues As Object = range.Cells.Value
            Dim obj As Object
            Dim y As Integer = 0
            Dim strArray(iMaxColumns) As String
            For Each obj In myvalues
                If y = 0 Then
                    If obj Is Nothing Then
                        Exit While
                    End If
                    If obj.ToString.Length = 0 Then
                        Exit While
                    End If
                End If
                If Not obj Is Nothing Then
                    If TypeOf obj Is Double And y = 0 Then
                        If strFormats.Length > y Then
                            If strFormats(y) = "Date" Then
                                If obj.ToString().Substring(0, 2) = "20" Then
                                    strArray(y) = obj.ToString()
                                Else
                                    Try
                                        strArray(y) = DateTime.FromOADate(CType(obj, Double)).ToShortDateString()
                                    Catch ex As Exception
                                        strArray(y) = obj.ToString()
                                    End Try
                                End If
                            Else
                                strArray(y) = obj.ToString()
                            End If
                        Else
                            strArray(y) = obj.ToString()
                        End If
                    Else
                        strArray(y) = obj.ToString()
                    End If
                End If
                y += 1
                If y = iMaxColumns Then Exit For
            Next
            A.Add(strArray)
            i += 1
            If i = iMaxLines Then Exit While
        End While
        xlApp.DisplayAlerts = False
        xlApp.ActiveWorkbook.Close()
        xlApp.Quit()
        NAR(oSheet)
        NAR(oBook)

        'Not sure about this but by not calling the Yes/No/Cancel dialog is avoided.
        'xlApp.Quit()
        NAR(xlApp)
        oSheet = Nothing
        oBook = Nothing
        xlApp = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Return A
    End Function
End Class
