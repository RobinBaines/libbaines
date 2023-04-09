'------------------------------------------------
'Name: Module XMLvb.
'Function: 
'Copyright Robin Baines 2007. All rights reserved.
'Created March 2007.
'Notes: 
'Modifications: 
'20161102 added overloeded OpenExcelBook so that an excel file can be created in any folder.
'20170906 Delete the xmlx file.
'Uses the last version of Excel Interop as the old one used in ExcelInterface 
'does not support OpenXML().
'The very few references to Excel therefore use an explicit reference to
'Imports Microsoft.Office.Interop.
'------------------------------------------------
'EXCEL AND REGIONAL SETTINGS.
'Excel stores data in the US format (CHECK or is it the land format of the Excel version).
'So 1234.50 is stored in the US Excel.
'When it displays data it shows it according to the regional settings.
'So in Dutch 1.234,50 in US 1,234.50
'If 1234,50 is stored in XML then US Excel interprets this as 123,450.00.
'When writing XML it is therefore important to write using the US format (CHECK if 
'using a US Excel). 
'A problem is that when running a user program on Dutch Regional settings any 
'double to string conversions will use ',' as decimal separator.
'For this reason I have converted ',' in 'number' strings to '.'.
'Check that this does not cause problems on Dutch Excel.
'See pWriteCell for the modification.
'20090518 RPB implemented blnFitToPage in CloseASheet
'20100105 RPB added BoldDecimal1 for RAP.
'20110309 RPB checked whether a Number type is really a number and changed to String with new format Normal10RightAligned.
'This allows format setting of N0 for a string and right aligns.
'20110504 RPB but previous modification causes problems where number is 0 length in forecasting. So test for this before checking whether 
'something is a number. May be necessary to test for colour in the Excel style.
'20110602 RPB Modified pWriteCell. Check whether a number is really a number before replacing , by . because System.Convert.ToDouble is sensitive to regional setting.
'20110602 RPB modified the number formats from "0" to "##,#0". this adds a ',' or in French ' ' as a 1000s separator. As this is the 
'normal regional setting format it is logical to do it in the Excel.
'20110817 RPB Modified SaveWB to save in xlsx format using code 51.
'Also modified extension of file from xls to xlsx in GetFilePath().
'20130224 Added strRowToRepeat so that the repeated row at the top of a printed page can be adjusted. Use by adding the row range: for example R1 or R2:R3.
'20200101 In CloseExcelBook tried to remove the Excel which remains behind after an excel sheet is closed by doing this but seems not to make a difference.
'20200914 Is Savewb() Added app.Visible and removed app.UserControl = True.
'------------------------------------------------
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data.SqlTypes
Imports System.Globalization
Imports Microsoft.VisualBasic.FileIO.FileSystem
Imports Microsoft.Office.Interop

Imports System.IO
Imports System.Xml
Imports System.Xml.XPath
Imports System.Xml.Xsl
Imports System.Diagnostics
Imports System.Reflection

Imports Microsoft.Office.Interop.Excel
Imports VBIDE = Microsoft.Vbe.Interop

Public Enum BorderWeight
    'Border weight of cells.
    None = 1
    Hairline
    Medium
    Thin
    Thick
End Enum
Public Enum Paths
    Local = 1
    Network = 2
End Enum
Public Enum ExcelStringFormats

    'The defined Formats. The user selects from this list.
    'The format of each shold be defined below. If it is not the first defined format is used.
    'If it is defined twice the first definition is used.
    Normal8
    Normal9
    Normal10
    NormalInt10
    NormalNr10
    NormalStringNr10    'Format is @ but Type is Number.
    NormalDate10

    Normal11
    Normal12
    Normal14
    Normal16

    Bold9
    Bold10
    Bold10RightAligned

    BoldInt10
    BoldNr10
    BoldDate10
    Bold11
    Bold12
    Bold16

    BoldRed10
    BoldNrRed10

    BoldUnderlined10
    NotVisible

    BoldBorder10
    BoldIntBorder10
    BoldNrBorder10
    BoldNrBorderRed10
    NormalBorder10

    NormalBorderRed10
    NormalNrBorder10
    NormalStringNrBorder10    'Format is @ but Type is Number.
    NormalIntBorder10
    NormalDateBorder10

    BoldEuroBorder10
    BoldEuroBorderRed10

    BoldBorderT12
    BoldBorderB12
    BoldBorderB11
    BoldBorderTB12
    BoldBorderBR12
    BoldEuroBorderTR10
    BoldEuroBorderTBR10

    NormalNrBackGroundGreen10
    NormalIntBackGroundGreen10

    Decimal0
    Decimal1
    Decimal2
    Decimal3
    Decimal4
    BoldDecimal0
    Currency2
    Percent0
    Percent2
    Percent4

    BoldPercent0
    NoStyle

    Footer

    '20100105 RPB added for RAP
    BoldDecimal1
    NormalRightAligned
    NormalWordWrap
End Enum


'Public Shared oExcel As Microsoft.Office.Interop.Excel.Application = Nothing
''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class XMLExcelCell
    Public strValue As String
    Public iStyle As ExcelStringFormats
    Public blnBold As Boolean
    Public Sub New(ByVal _strValue As String, ByVal _iStyle As ExcelStringFormats)
        strValue = _strValue
        iStyle = _iStyle
        blnBold = False
    End Sub
End Class

''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class XMLExcelInterface

    '20110423 RPB Is set in OpenExcelBook and if true is used to adjust styles for word wrapping. 
    Dim blnWordWrap As Boolean = False

    Dim strPrtArea As String

    '20111118 Added these to allow users to change the True to a False.
    Dim _strFalse As String = "False"
    Property strFalse() As String
        Get
            Return _strFalse
        End Get
        Set(ByVal value As String)
            _strFalse = value
        End Set
    End Property

    Dim _strTrue As String = "True"
    Property strTrue() As String
        Get
            Return _strTrue
        End Get
        Set(ByVal value As String)
            _strTrue = value
        End Set
    End Property

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
    Dim originalCulture As System.Globalization.CultureInfo
    Dim strFirstSheet As String
    Dim strLocalOutput = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    Dim strNetworkOutput = My.Computer.FileSystem.SpecialDirectories.MyDocuments    'My.Settings.OutputPath
    Dim stW As StreamWriter
    Dim strFileName As String
    Dim strOriginalFileName As String
    Dim strCurrentSheetName As String
    Dim blnIsFirstSheet As Boolean
    Dim iColourIndex As Integer

    Public ExcelFormats As New List(Of ExcelFormat)

    Dim XMLData As New List(Of String)
    Dim blnLandscape As Boolean 'Store parameter between Open and closing a sheet.
    Dim strFooter As String
    Dim strHeaderText As String
    Dim blnShowHeader As Boolean
    Dim dLeftMargin As Double
    Dim dRightMargin As Double
    Dim dTopMargin As Double
    Dim dBottomMargin As Double
    Dim dHeaderMargin As Double
    Dim dFooterMargin As Double
    Dim iPaperSize As Double
    Dim blnFitToPage As Boolean
    Dim blnRepeatTopRow As Boolean
    Dim iPrintQuality As Integer
    Dim iZoom As Integer
    Dim strSeparator As String
    Dim str1000Separator As String
    Dim strRowToRepeat As String    'R1 or for example R2:R3.

    Public Sub New(ByVal strOutput As String)

        'Defines the network destination for files.
        strNetworkOutput = strOutput
        XMLData.Clear()
        Dim value As NumberFormatInfo
        value = NumberFormatInfo.CurrentInfo
        strSeparator = value.NumberDecimalSeparator()
        str1000Separator = value.NumberGroupSeparator()

        'CreateFormat("#" + str1000Separator + "##0" + strSeparator + "000")
        CreateFormat("#,##0.000")
    End Sub

    Public Sub New(ByVal strOutput As String, ByVal strDecimalFormat As String)

        'Added strDecimalFormat parameter to allow control over the digits behind decimal point.
        'Defines the network destination for files.

        strNetworkOutput = strOutput
        XMLData.Clear()
        Dim value As NumberFormatInfo
        value = NumberFormatInfo.CurrentInfo
        strSeparator = value.NumberDecimalSeparator()
        str1000Separator = value.NumberGroupSeparator()

        CreateFormat(strDecimalFormat)
    End Sub

    Public Sub New(ByVal strOutput As String, ByVal blnSetBorders As Boolean)

        'Set all borders to hairline if true.
        Me.New(strOutput)
        For Each f As ExcelFormat In ExcelFormats
            f.BorderBottom = True
            f.BorderLeft = True
            f.BorderTop = True
            f.BorderRight = True
            f.BorderWeight = BorderWeight.Hairline
        Next
    End Sub

    Private Function strGetStyle(ByVal iStyleName As ExcelStringFormats) As String

        'Lookup the first Format which matches the Style and return the StyleName.
        Dim strRet As String
        Dim ef As ExcelFormat
        strRet = ""
        If iStyleName <> ExcelStringFormats.NoStyle Then
            For Each ef In ExcelFormats
                If strRet = "" Then strRet = ef.Name
                If ef.iName = iStyleName Then
                    strRet = ef.Name
                    Exit For
                End If
            Next ef
        End If

        Return strRet
    End Function

    Private Function strGetType(ByVal iStyleName As ExcelStringFormats) As String

        'Lookup the first Format which matches the Style and return the StyleType (String or Number etc).
        Dim strRet As String
        Dim ef As ExcelFormat
        strRet = ""
        For Each ef In ExcelFormats
            If strRet = "" Then strRet = ef.Type
            If ef.iName = iStyleName Then
                strRet = ef.Type
                Exit For
            End If
        Next ef
        Return strRet
    End Function

    Private Sub CreateFormat(ByVal strDecimalFormat As String)
        Dim value As NumberFormatInfo
        value = NumberFormatInfo.CurrentInfo

        ExcelFormats.Clear()
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Normal8, "Normal8", False, False, False, "@", "String", 8, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Normal9, "Normal9", False, False, False, "@", "String", 9, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Normal10, "Normal10", False, False, False, "@", "String", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.NormalInt10, "NormalInt10", False, False, False, "0", "Number", 10, "Swiss"))

        ' Feb 2008 ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.NormalNr10, "NormalNr10", False, "#,##0", "Number", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.NormalNr10, "NormalNr10", False, False, False, strDecimalFormat, "Number", 10, "Swiss"))

        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.NormalStringNr10, "NormalStringNr10", False, False, False, "@", "Number", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.NormalDate10, "NormalDate10", False, False, False, "d\-mm\-yy;@", "DateTime", 10, "Swiss"))

        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Normal11, "Normal11", False, False, False, "@", "String", 11, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Normal12, "Normal12", False, False, False, "@", "String", 12, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Normal14, "Normal14", False, False, False, "@", "String", 14, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Normal16, "Normal16", False, False, False, "@", "String", 16, "Swiss"))

        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Bold9, "Bold9", True, False, False, "@", "String", 9, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Bold10, "Bold10", True, False, False, "@", "String", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Bold10RightAligned, "Bold10RightAligned", True, False, True, "@", "String", 10, "Swiss"))

        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldInt10, "BoldInt10", True, False, False, "0", "Number", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldNr10, "BoldNr10", True, False, False, strDecimalFormat, "Number", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldDate10, "BoldDate10", True, False, False, "d\-mm\-yy;@", "DateTime", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Bold11, "Bold11", True, False, False, "@", "String", 11, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Bold12, "Bold12", True, False, False, "@", "String", 12, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Bold16, "Bold16", True, False, False, "@", "String", 16, "Swiss"))

        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldRed10, "BoldRed10", False, "#FF0000", False, False, False, False, BorderWeight.None, "@", "String", 10, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldNrRed10, "BoldNrRed10", False, "#FF0000", False, False, False, False, BorderWeight.None, strDecimalFormat, "Number", 10, "", "Swiss", ""))

        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldUnderlined10, "BoldUnderlined10", True, "", False, False, False, False, BorderWeight.None, "@", "String", 10, "Single", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.NotVisible, "NotVisible", False, "#FFFFFF", False, False, False, False, BorderWeight.None, "@", "String", 9, "", "Swiss", ""))

        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldBorder10, "BoldBorder10", True, "", True, True, True, True, BorderWeight.Hairline, "@", "String", 10, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldIntBorder10, "BoldIntBorder10", True, "", True, True, True, True, BorderWeight.Hairline, "##", "Number", 10, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldNrBorder10, "BoldNrBorder10", True, "", True, True, True, True, BorderWeight.Hairline, strDecimalFormat, "Number", 10, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldNrBorderRed10, "BoldNrBorderRed10", True, "#FF0000", True, True, True, True, BorderWeight.Hairline, strDecimalFormat, "Number", 10, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.NormalBorder10, "NormalBorder10", False, "", True, True, True, True, BorderWeight.Hairline, "@", "String", 10, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.NormalBorderRed10, "NormalBorderRed10", False, "#FF0000", True, True, True, True, BorderWeight.Hairline, "@", "String", 10, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.NormalIntBorder10, "NormalIntBorder10", False, "", True, True, True, True, BorderWeight.Hairline, "0", "Number", 10, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.NormalNrBorder10, "NormalNrBorder10", False, "", True, True, True, True, BorderWeight.Hairline, strDecimalFormat, "Number", 10, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.NormalStringNrBorder10, "NormalStringNrBorder10", False, "", True, True, True, True, BorderWeight.Hairline, "@", "Number", 10, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.NormalDateBorder10, "NormalDateBorder10", False, "", True, True, True, True, BorderWeight.Hairline, "d\-mm\-yy;@", "DateTime", 10, "", "Swiss", ""))

        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldEuroBorder10, "BoldEuroBorder10", True, "", True, True, True, True, BorderWeight.Hairline, "\€#,##0.00", "Number", 10, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldEuroBorderRed10, "BoldEuroBorderRed10", True, "#FF0000", True, True, True, True, BorderWeight.Hairline, "\€#,##0.00", "Number", 10, "", "Swiss", ""))

        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldBorderT12, "BoldBorderT12", True, "", True, False, False, False, BorderWeight.Hairline, "@", "String", 12, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldBorderB12, "BoldBorderB12", True, "", False, False, True, False, BorderWeight.Hairline, "@", "String", 12, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldBorderB11, "BoldBorderB11", True, "", False, False, True, False, BorderWeight.Hairline, "@", "String", 11, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldBorderTB12, "BoldBorderTB12", True, "", True, False, True, False, BorderWeight.Hairline, "@", "String", 12, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldBorderBR12, "BoldBorderBR12", True, "", False, False, True, True, BorderWeight.Hairline, "@", "String", 12, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldEuroBorderTR10, "BoldEuroBorderTR10", True, "", True, False, False, True, BorderWeight.Hairline, "\€#,##0.00", "Number", 10, "", "Swiss", ""))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldEuroBorderTBR10, "BoldEuroBorderTBR10", True, "", True, False, True, True, BorderWeight.Hairline, "\€#,##0.00", "Number", 10, "", "Swiss", ""))

        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.NormalNrBackGroundGreen10, "NormalNrBackGroundGreen10", False, "", False, False, False, False, BorderWeight.None, strDecimalFormat, "Number", 10, "", "Swiss", "#99CC00"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.NormalIntBackGroundGreen10, "NormalIntBackGroundGreen10", False, "", False, False, False, False, BorderWeight.None, 0, "Number", 10, "", "Swiss", "#99CC00"))

        '20110602 RPB modified the number formats from "0" to "##,#0". this adds a ',' or in French ' ' as a 1000s separator. As this is the 
        'normal regional setting format it is logical to do it in the Excel.
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Decimal0, "Decimal0", False, False, False, "##,#0", "Number", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Decimal1, "Decimal1", False, False, False, "##,#0.0", "Number", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Decimal2, "Decimal2", False, False, False, "##,#0.00", "Number", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Decimal3, "Decimal3", False, False, False, "##,#0.000", "Number", 10, "Swiss"))
        'ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Decimal3, "Decimal3", False, False, False, "0.000", "Number", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Decimal4, "Decimal4", False, False, False, "##,#0.0000", "Number", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Currency2, "Currency2", False, False, False, "\€#,##0.00", "Number", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldDecimal0, "BoldDecimal0", True, False, False, "##,#0", "Number", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldDecimal1, "BoldDecimal1", True, False, False, "##,#0.0", "Number", 10, "Swiss"))


        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Percent0, "Percent0", False, False, False, "0%", "Number", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Percent2, "Percent2", False, False, False, "0.00%", "Number", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Percent4, "Percent4", False, False, False, "0.0000%", "Number", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.BoldPercent0, "BoldPercent0", True, False, False, "0%", "Number", 10, "Swiss"))

        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.Footer, "Footer", False, True, False, "@", "String", 8, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.NormalRightAligned, "NormalRightAligned", False, False, True, "@", "String", 10, "Swiss"))
        ExcelFormats.Add(New ExcelFormat(ExcelStringFormats.NormalWordWrap, "NormalWordWrap", False, False, False, "@", "String", 10, "Swiss", True))

    End Sub

    Private Function GetBorderWeight(ByVal i As BorderWeight) As XlBorderWeight

        'Convert ExcelInterface public enum to Excel value.
        If i = BorderWeight.Hairline Then Return XlBorderWeight.xlHairline
        If i = BorderWeight.Medium Then Return XlBorderWeight.xlMedium
        If i = BorderWeight.Thin Then Return XlBorderWeight.xlThin
        If i = BorderWeight.Thick Then Return XlBorderWeight.xlThick
        Return XlBorderWeight.xlHairline
    End Function

#Region "FileOpen"
    Private Function GetFilePath(ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal strFileName As String) As String

        'Create path to saved book.
        Return GetDirectory(pPath, strSubDirectory) & strFileName & ".xlsx"
    End Function

    Private Function GetDirectory(ByVal pPath As Paths, ByVal strSubDirectory As String) As String

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

    Private Sub CreateStyles()

        'Write all the Styles to the XML file.
        stW.WriteLine("<Styles>")
        Dim ef As ExcelFormat

        For Each ef In ExcelFormats
            stW.WriteLine("<Style ss:ID='" & ef.Name & "'>")

            If ef.BorderBottom = True Or ef.BorderLeft = True Or ef.BorderRight = True Or ef.BorderTop = True Then

                stW.WriteLine("<Borders>")
                If ef.BorderBottom = True Then
                    stW.WriteLine("<Border ss:Position='Bottom' ss:LineStyle='Continuous' ss:Weight='" & _
                    GetBorderWeight(ef.BorderWeight) & "'/>")
                End If
                If ef.BorderLeft = True Then
                    stW.WriteLine("<Border ss:Position='Left' ss:LineStyle='Continuous' ss:Weight='" & _
                    GetBorderWeight(ef.BorderWeight) & "'/>")
                End If
                If ef.BorderRight = True Then
                    stW.WriteLine("<Border ss:Position='Right' ss:LineStyle='Continuous' ss:Weight='" & _
                    GetBorderWeight(ef.BorderWeight) & "'/>")
                End If
                If ef.BorderTop = True Then
                    stW.WriteLine("<Border ss:Position='Top' ss:LineStyle='Continuous' ss:Weight='" & _
                    GetBorderWeight(ef.BorderWeight) & "'/>")
                End If
                stW.WriteLine("</Borders>")
            End If
            Dim str As String
            str = "<Font x:Family='" & ef.Font & "'  ss:Size='" & ef.FontSize.ToString & "' "
            If ef.Colour.Length <> 0 Then
                str = str & " ss:Color='" & ef.Colour & "'"
            End If
            If ef.Bold = True Then
                str = str & " ss:Bold='1'"
            End If
            If ef.Italic = True Then
                str = str & " ss:Italic='1'"
            End If

            If ef.Underlined <> "" Then
                str = str & " ss:Underline='" & ef.Underlined & "'"
            End If

            str = str & " />"
            stW.WriteLine(str)
            If ef.BackGround <> "" Then
                stW.WriteLine("<Interior ss:Color='" & ef.BackGround & "' ss:Pattern='Solid'/>")
            End If

            '20111116
            If ef.Format = "@" Then
                stW.WriteLine("<NumberFormat/>")
            Else
                stW.WriteLine("<NumberFormat ss:Format='" & ef.Format & "'/>")
            End If

            If ef.RightAligned = True Then
                stW.WriteLine("<Alignment ss:Horizontal='Right' ss:Vertical='Bottom' ss:WrapText='1'/>")
            Else
                'If ef.WordWrap = True Then
                '20110423 RPB. Add ss:WrapText='1' word wrap text. 
                'See also modification of 20110423 to the Row attribute.
                If blnWordWrap = True Or ef.WordWrap = True Then
                    stW.WriteLine("<Alignment ss:Vertical='Bottom' ss:WrapText='1' />")
                Else
                    stW.WriteLine("<Alignment ss:Vertical='Bottom' />")
                End If
                'Else
                'stW.WriteLine("<Alignment ss:Vertical='Bottom' />")
                'End If
            End If
            stW.WriteLine("</Style>")
        Next

        'Add test sysles here.
        'stW.WriteLine("<Style ss:ID='s999'>")
        'stW.WriteLine("<Alignment ss:Vertical='Bottom' ss:WrapText='1'/>")
        'stW.WriteLine("</Style>")

        stW.WriteLine("</Styles>")
    End Sub

    Public Sub OpenExcelBook(ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal _strFileName As String, _
       ByVal blnUnique As Boolean, ByVal strTemplatePath As String)
        OpenExcelBook(pPath, strSubDirectory, _strFileName, blnUnique, strTemplatePath, False)
    End Sub

    Public Sub OpenExcelBook(ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal _strFileName As String, _
        ByVal blnUnique As Boolean, ByVal strTemplatePath As String, ByVal _blnWordWrap As Boolean)

        blnWordWrap = _blnWordWrap
        strFileName = _strFileName

        '20121121 Add the start up path because Terminal Server does not have that as the default path.
        strTemplatePath = System.Windows.Forms.Application.StartupPath + "\" + strTemplatePath
        If strFileName.Length <> 0 Then

            'Create target directory if it does not exist already.
            If Not My.Computer.FileSystem.DirectoryExists(GetDirectory(pPath, strSubDirectory)) Then
                'Create target directory if it does not exist already.
                My.Computer.FileSystem.CreateDirectory(GetDirectory(pPath, strSubDirectory))
            End If

            'Try to save the Excel file with the new name.
            strFileName = MakeUniquePath(0, blnUnique, pPath, strSubDirectory, _strFileName)
            strOriginalFileName = strFileName
            strFileName = strFileName.Replace(".xls", ".xml")

            FileCopy(strTemplatePath, strFileName)
            stW = New StreamWriter(strFileName, True)
            blnIsFirstSheet = True
            CreateStyles()
            strFirstSheet = ""
            strCurrentSheetName = ""
            iColourIndex = 0
        End If

    End Sub

    '20161102 added overloeded OpenExcelBook so that an excel file can be created in any folder.
    Public Sub OpenExcelBook(ByVal strPath As String, ByVal _strFileName As String, _
    ByVal blnUnique As Boolean, ByVal strTemplatePath As String, ByVal _blnWordWrap As Boolean)

        blnWordWrap = _blnWordWrap
        strFileName = _strFileName

        '20121121 Add the start up path because Terminal Server does not have that as the default path.
        strTemplatePath = System.Windows.Forms.Application.StartupPath + "\" + strTemplatePath
        If strFileName.Length <> 0 Then

            If Not strPath.EndsWith("\") Then
                strPath = strPath + "\"
            End If

            'Create target directory if it does not exist already.
            If Not My.Computer.FileSystem.DirectoryExists(strPath) Then
                'Create target directory if it does not exist already.
                My.Computer.FileSystem.CreateDirectory(strPath)
            End If

            'In the above version of OpenExcelBook the MakeUniquePath makes a unique name and then calls GetDirectory to create the complete name.
            'Here the user is prompted to write over any other version of the xlsx.
            'strFileName = MakeUniquePath(0, blnUnique, pPath, strSubDirectory, _strFileName)
            'GetDirectory(pPath, strSubDirectory) & strFileName & ".xlsx"
            strFileName = strPath + _strFileName & ".xlsx"
            strOriginalFileName = strFileName
            strFileName = strFileName.Replace(".xls", ".xml")

            FileCopy(strTemplatePath, strFileName)
            stW = New StreamWriter(strFileName, True)
            blnIsFirstSheet = True
            CreateStyles()
            strFirstSheet = ""
            strCurrentSheetName = ""
            iColourIndex = 0
        End If
    End Sub

    Public Sub OpenExcelBookWithString(ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal _strFileName As String, _
        ByVal blnUnique As Boolean, ByVal strTemplate As String)

        strFileName = _strFileName

        If strFileName.Length <> 0 Then

            'Create target directory if it does not exist already.
            If Not My.Computer.FileSystem.DirectoryExists(GetDirectory(pPath, strSubDirectory)) Then
                My.Computer.FileSystem.CreateDirectory(GetDirectory(pPath, strSubDirectory))
            End If

            'Try to save the Excel file with the new name.
            strFileName = MakeUniquePath(0, blnUnique, pPath, strSubDirectory, _strFileName)
            strOriginalFileName = strFileName
            strFileName = strFileName.Replace(".xls", ".xml")

            'FileCopy(strTemplatePath, strFileName)
            stW = New StreamWriter(strFileName, False)
            stW.WriteLine(strTemplate)
            blnIsFirstSheet = True
            CreateStyles()
            strFirstSheet = ""
            strCurrentSheetName = ""
            iColourIndex = 0
        End If
    End Sub

    Public Sub OpenExcelBook(ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal strFileName As String, _
        ByVal strHeaderText As String, ByVal strFooter As String, ByVal blnLandscape As Boolean, ByVal strTemplatePath As String)
        ',         ByVal ExcelFormats As List(Of ExcelFormat))
        OpenExcelBook(pPath, strSubDirectory, strFileName, False, strTemplatePath)
        NewSheet(strHeaderText, strHeaderText, blnLandscape, strFooter)
    End Sub

    Public Sub OpenExcelBook(ByVal pPath As Paths, ByVal strSubDirectory As String, ByVal strFileName As String, _
        ByVal strHeaderText As String, ByVal strFooter As String, ByVal blnLandscape As Boolean, ByVal strTemplatePath As String, _
        ByVal _blnWordWrap As Boolean)
        ',         ByVal ExcelFormats As List(Of ExcelFormat))
        OpenExcelBook(pPath, strSubDirectory, strFileName, False, strTemplatePath, _blnWordWrap)
        NewSheet(strHeaderText, strHeaderText, blnLandscape, strFooter)
    End Sub

    Public Function SetActivePrinter(ByVal strPrinter As String)
        ', ByVal iPaperSize As Integer, ByVal iPrintQuality As Integer) As Boolean
        Dim blnRet As Boolean
        blnRet = True
        'Try
        '    xlApp.Application.ActivePrinter = strPrinter
        'Catch ex As Exception
        '    MsgBox("Could not select Printer: " & strPrinter & " " & ex.Message, MsgBoxStyle.OkOnly)
        '    blnRet = False
        'Finally
        'End Try
        Return blnRet
    End Function
#End Region

#Region "SheetOpenClose"
    Private Sub CloseASheet()

        If Not stW Is Nothing Then

            'Write the XML to open a sheet.
            stW.WriteLine("<Worksheet ss:Name='" & strCurrentSheetName & "'>")
            stW.WriteLine("<Names>")

            If RowCount = 0 And ColumnCount = 0 Then
                stW.WriteLine("<NamedRange ss:Name='Print_Area' ss:RefersTo=""='" & strCurrentSheetName & "'!R1C1:R41C14""/>")
            Else
                stW.WriteLine("<NamedRange ss:Name='Print_Area' ss:RefersTo=""='" & strCurrentSheetName & "'!R1C1:R" & _
                Me.RowCount.ToString() & "C" & (Me.ColumnCount - 1).ToString() & """/>")
            End If

            '20130224 Allow adjustment of Repeat Top Row.
            If strRowToRepeat.Length > 0 Then stW.WriteLine("<NamedRange ss:Name='Print_Titles' ss:RefersTo=""='" & strCurrentSheetName & "'!" & strRowToRepeat & """/>")
            '& "'!R1""/>")


            stW.WriteLine("</Names>")
            stW.WriteLine("<Table>")

            'Write the table.
            WriteXMLData()

            'Write the XML to close a sheet.
            stW.WriteLine("</Table>")
            stW.WriteLine("<WorksheetOptions xmlns='urn:schemas-microsoft-com:office:excel'>")
            stW.WriteLine("<PageSetup>")
            If blnLandscape = True Then
                stW.WriteLine("<Layout x:Orientation='Landscape'/>")
            Else
            End If

            stW.WriteLine("<Header x:Margin='0'")
            If Not strHeaderText Is Nothing Then
                'stW.WriteLine("x:Data='&amp;LUsing data from " & strCurrentSheetName & ".&amp;CPrinted on &amp;D &amp;T. &amp;RPage &amp;P of &amp;N'/>")
                stW.WriteLine("x:Data='" & strHeaderText & "'/>")
            Else
                stW.WriteLine("x:Data=''/>")
            End If

            If Not strFooter Is Nothing Then
                stW.WriteLine("<Footer x:Margin='0' x:Data='" & strFooter & "'/>")
            Else
                stW.WriteLine("<Footer x:Margin='0' x:Data=''/>")
            End If

            '2009090 RPB added CultureInfo to force writing in us format.
            Dim MyCultureInfo As CultureInfo = New CultureInfo("en-US")
            stW.WriteLine("<PageMargins x:Bottom='" & dBottomMargin.ToString("0.00", MyCultureInfo) & "' x:Left='" & _
                dLeftMargin.ToString("0.00", MyCultureInfo) & "' x:Right='" & dRightMargin.ToString("0.00", MyCultureInfo) & "' x:Top='" & _
                dTopMargin.ToString("0.00", MyCultureInfo) & "'/>")

            stW.WriteLine("</PageSetup>")
            stW.WriteLine("<FitToPage/>")
            stW.WriteLine("<Print>")

            '20090518 blnFitToPage is set by SetFitTo1Page
            If blnFitToPage = False Then
                stW.WriteLine("<FitHeight>0</FitHeight>")
            End If

            stW.WriteLine("<ValidPrinterInfo/>")
            stW.WriteLine("<PaperSizeIndex>9</PaperSizeIndex>")
            stW.WriteLine("<Scale>63</Scale>")
            stW.WriteLine("<HorizontalResolution>-3</HorizontalResolution>")
            stW.WriteLine("<VerticalResolution>0</VerticalResolution>")
            stW.WriteLine("</Print>")
            stW.WriteLine("<Selected/>")
            stW.WriteLine("<ProtectObjects>False</ProtectObjects>")
            stW.WriteLine("<ProtectScenarios>False</ProtectScenarios>")
            stW.WriteLine("</WorksheetOptions>")
            stW.WriteLine("</Worksheet>")

        End If
    End Sub

    '20130224 Allow adjustment of Repeat Top Row.
    Public Sub NewSheet(ByVal _strSheetName As String, _
                        ByVal _strHeaderText As String, _
                        ByVal _blnLandScape As Boolean, _
                        ByVal _strFooter As String)

        NewSheet(_strSheetName, _strHeaderText, _blnLandScape, _strFooter, "R1")

    End Sub

    '20130224 Allow adjustment of Repeat Top Row.
    Public Sub NewSheet(ByVal _strSheetName As String, _
                      ByVal _strHeaderText As String, _
                      ByVal _blnLandScape As Boolean, _
                      ByVal _strFooter As String, _strRowToRepeat As String)

        'Create a new work sheet.
        If (blnIsFirstSheet = False) Then
            CloseASheet()
            SetRow(0)
            ColumnCount = 0
            XMLData.Clear()
        Else
            blnIsFirstSheet = False
        End If
        blnLandscape = _blnLandScape
        strFooter = _strFooter
        strCurrentSheetName = _strSheetName
        strHeaderText = _strHeaderText
        dBottomMargin = 0.58
        dLeftMargin = 0.89
        dRightMargin = 0.56
        dTopMargin = 0.58

        'R1 or for example R2:R3.
        strRowToRepeat = _strRowToRepeat

    End Sub

    Public Sub NewSheet(ByVal _strSheetName As String, _
                ByVal _strHeaderText As String, _
                ByVal _blnLandScape As Boolean, _
                ByVal _strFooter As String, _
                ByVal _blnShowHeader As Boolean, _
                ByVal _dLeftMargin As Double, ByVal _dRightMargin As Double, ByVal _dTopMargin As Double, _
                ByVal _dBottomMargin As Double, _
                ByVal _dHeaderMargin As Double, _
                ByVal _dFooterMargin As Double, ByVal _iPaperSize As Double, ByVal _blnFitToPage As Boolean, _
                ByVal _blnRepeatTopRow As Boolean, ByVal _iPrintQuality As Integer, ByVal _iZoom As Integer)

        NewSheet(_strSheetName, _strHeaderText, _blnLandScape, _strFooter, "R1")
        blnShowHeader = _blnShowHeader
        dLeftMargin = _dLeftMargin
        dRightMargin = _dRightMargin
        dTopMargin = _dTopMargin
        dBottomMargin = _dBottomMargin
        dHeaderMargin = _dHeaderMargin
        dFooterMargin = _dFooterMargin
        iPaperSize = _iPaperSize
        blnFitToPage = _blnFitToPage
        blnRepeatTopRow = _blnRepeatTopRow
        iPrintQuality = _iPrintQuality
        iZoom = _iZoom

    End Sub

    Private Sub StoreXMLData(ByVal strT As String)
        XMLData.Add(strT)
    End Sub

    Private Sub WriteXMLData()
        Dim strT As String
        For Each strT In XMLData
            stW.WriteLine(strT)

            '20200102
            strT = ""
        Next
    End Sub

    Public Sub strCreateExcelSheet(ByVal dg As DataGridView, ByVal strFileName As String, ByVal strHeader As String, ByVal strFoot As String, _
    ByVal strTagFilter As String, ByVal strXMLTemplate As String)
        OpenExcelBook(Paths.Local, "", strFileName, strHeader, strFoot, blnLandscape, strXMLTemplate)

        WriteDataGrid(dg, strTagFilter, False, 0, False)
        CloseExcelBook()

    End Sub

    Public Sub PrintMasterChild(ByVal dgMaster As DataGridView, ByVal dgChild As DataGridView, _
    ByVal strSubDirectory As String, ByVal strFileName As String, ByVal strHeader As String, ByVal strFoot As String, _
    ByVal strTagFilter As String, ByVal strXMLTemplate As String)

        OpenExcelBook(Paths.Local, "", strFileName, strHeader, strFoot, blnLandscape, strXMLTemplate)
        WriteDataGrid(dgMaster, strTagFilter, True, 0, False)
        WriteStringToExcel(" ", ExcelStringFormats.Normal10)
        WriteDataGrid(dgChild, strTagFilter, False, 0, True, False, True)
        CloseExcelBook()
    End Sub
  
#End Region
#Region "Close"
    Dim wb As Workbook
    Dim app As Microsoft.Office.Interop.Excel.Application
    Public Sub SaveWb()
        If Not wb Is Nothing Then

            'The saveas converts the xml to an xls file.
            wb.SaveAs(strOriginalFileName, 51)
            Try
                '20200914 Added app.Visible and removed app.UserControl = True.
                app = wb.Application
                app.Visible = True
                'app.UserControl = True

            Catch ex As Exception
                'wb.Parent.UserControl = True
            End Try

            '20200102
            NAR(wb)

            '20170906 Delete the xmlx file.
            Dim strFName As String = strOriginalFileName.Replace(".xlsx", ".xmlx")
            If FileExists(strFName) = True Then
                My.Computer.FileSystem.DeleteFile(strFName)
            End If
        End If

    End Sub

    ''' <summary>
    ''' 20090909 Added flag to prevent conversion to us culture.
    ''' </summary>
    ''' <param name="blnRetainCulture"></param>
    ''' <remarks></remarks>
    Public Sub CloseExcelBook(ByVal blnRetainCulture As Boolean)
        CloseExcelBook(False, blnRetainCulture)
    End Sub

    Public Sub CloseExcelBook()
        CloseExcelBook(True, False)
    End Sub

    Public Function CloseExcelBook(ByVal blnCompleteProcessing As Boolean, ByVal blnRetainCulture As Boolean) As String

        'Write the XML required to finish the XML file.
        CloseExcelBookAndExitExcel()

        'Open the XML file in Excel.
        'Dim oExcel As New Microsoft.Office.Interop.Excel.Application
        'oExcel = New Microsoft.Office.Interop.Excel.Application
        Dim oExcel = CreateObject("Excel.Application")

        oExcel.Visible = True
        oExcel.UserControl = True

        ' ''20200101 In CloseExcelBook tried to remove the Excel which remains behind after an excel sheet is closed by doing this but seems not to make a difference.
        'If oExcel.COMAddIns.Count > 0 Then

        '    For i As Integer = 0 To oExcel.COMAddIns.Count - 1
        '        Try
        '            oExcel.COMAddIns.Item(i).Connect = False
        '            'MsgBox("Count " + i.ToString())
        '            'The following is the Microsoft Power Pivot addin for Excel.
        '            'MsgBox("App name " + oExcel.COMAddIns.Item(i).Description)
        '        Catch ex As Exception
        '            'MsgBox("Exception Count " + i.ToString() + " " + oExcel.COMAddIns.Item(i).Description + " " + oExcel.COMAddIns.Item(i).Application)
        '        End Try

        '    Next
        'End If

        'Dim objWMIService
        'Dim colProcessList
        'Dim objProcess
        'Dim strComputer

        'strComputer = "."
        'Try
        '    objWMIService = GetObject("winmgmts://./root/cimv2")  ' Task mgr
        '    If Not objWMIService Is Nothing Then
        '        Try
        '            colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name in ('EXCEL.EXE') ")  '''''','Chrome.exe','iexplore.exe'
        '            For Each objProcess In colProcessList
        '                MsgBox(objProcess.ToString())
        '                objProcess.Terminate()
        '                MsgBox("2." + objProcess.ToString())
        '            Next
        '        Catch ex As Exception
        '            MsgBox(ex.Message)
        '        End Try
        '    Else
        '        MsgBox("winmgmts://./root/cimv2 NOT FOUND")
        '    End If
        'Catch ex As Exception
        '    MsgBox("GetObject" + ex.Message)
        'End Try


        Dim oldCI As System.Globalization.CultureInfo = _
            System.Threading.Thread.CurrentThread.CurrentCulture

        If blnRetainCulture = False Then
            System.Threading.Thread.CurrentThread.CurrentCulture = _
                New System.Globalization.CultureInfo("en-US")
        End If
        Dim oBooks As Workbooks
        oBooks = oExcel.Workbooks
        wb = oBooks.Open(strFileName)    '.OpenXML(strFileName)

        'Dim oSheet As Worksheet = wb.Application.ActiveWorkbook.ActiveSheet
        'oSheet.Copy()
        'oSheet.Select()
        'Clipboard.SetDataObject(oSheet, True)

        'Save as and delete xml file.
        If blnCompleteProcessing = True Then
            SaveWb()
        End If
        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        'Quit closes the excel workbook but does not remove the background process.
        'oExcel.Quit()

        '20200101 Added this.
        GC.Collect()
        GC.WaitForPendingFinalizers()

        NAR(oBooks)
        NAR(oExcel)

        GC.Collect()
        GC.WaitForPendingFinalizers()
        'System.Threading.Thread.CurrentThread.CurrentCulture = Nothing

        Return strFileName
    End Function

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

    Public Function CloseExcelBookAndExitExcel() As String

        'Write the XML required to finish the XML file.
        CloseASheet()
        stW.WriteLine("</Workbook>")
        stW.Flush()
        stW.Close()
        stW.Dispose()
        stW = Nothing
        XMLData.Clear()

        '20200204
        XMLData = Nothing

        Return strFileName
    End Function

#End Region
#Region "DataGrid"

    Public Sub WriteColumnWidths(ByVal dg As DataGridView, ByVal strTagFilter As String, ByVal blnCurrentBold As Boolean, _
        ByVal iFirstColumn As Integer)

        'Display DataGridView in Excel.
        'Function requires that all headertexts of the datagridview are unique.
        'Dim Cells As List(Of XMLExcelCell)
        Dim ColumnWidth As List(Of Int32)
        'Cells = New List(Of XMLExcelCell)(dg.Columns.Count)
        ColumnWidth = New List(Of Int32)(dg.Columns.Count)
        ColumnWidth.Clear()
        'Cells.Clear()
        Dim iCellsColumn = 0

        'Set up column widths, formats and headers.
        For Each iColumn As DataGridViewColumn In dg.Columns
            If CheckTag(iColumn, strTagFilter) = True Then
                '       Cells.Add(New XMLExcelCell("", ExcelStringFormats.Bold10))
                '      Cells(iCellsColumn).strValue = iColumn.HeaderText()
                ColumnWidth.Add(iColumn.Width)
                iCellsColumn = iCellsColumn + 1
            End If
        Next
        SetAutofit(ColumnWidth)

        'Set formats for data.
        'Cells.Clear()

    End Sub

    Public Sub WriteDataGrid(ByVal dg As DataGridView, ByVal strTagFilter As String, ByVal blnCurrentBold As Boolean, _
        ByVal iFirstColumn As Integer, ByVal blnAutofit As Boolean)
        WriteDataGrid(dg, strTagFilter, blnCurrentBold, iFirstColumn, blnAutofit, True, True)
    End Sub

    Public Sub WriteDataGrid(ByVal dg As DataGridView, ByVal strTagFilter As String, ByVal blnCurrentBold As Boolean, _
        ByVal iFirstColumn As Integer, ByVal blnAutofit As Boolean, ByVal blnWriteColumnWidths As Boolean, _
        ByVal blnWriteHeaders As Boolean)

        'Display DataGridView in Excel.
        'Function requires that all headertexts of the datagridview are unique.
        Dim Cells As List(Of XMLExcelCell)
        Dim ColumnWidth As List(Of Int32)
        Cells = New List(Of XMLExcelCell)(dg.Columns.Count)
        ColumnWidth = New List(Of Int32)(dg.Columns.Count)
        ColumnWidth.Clear()
        Cells.Clear()
        Dim iCellsColumn = 0

        'Set up column widths, formats and headers.
        For Each iColumn As DataGridViewColumn In dg.Columns
            If CheckTag(iColumn, strTagFilter) = True Then
                If iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("N") Or iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("P") _
                    Or iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("C") Then
                    Cells.Add(New XMLExcelCell("", ExcelStringFormats.Bold10RightAligned))
                Else
                    Cells.Add(New XMLExcelCell("", ExcelStringFormats.Bold10))
                End If
                Cells(iCellsColumn).strValue = iColumn.HeaderText()
                ColumnWidth.Add(iColumn.Width)
                iCellsColumn = iCellsColumn + 1
            End If
        Next
        If blnWriteColumnWidths = True Then
            SetAutofit(ColumnWidth)
        End If
        If blnWriteHeaders = True Then
            WriteStringToExcel(Cells, 1, False, ExcelStringFormats.Bold10)
        End If

        'Set formats for data.
        Cells.Clear()
        For Each iColumn As DataGridViewColumn In dg.Columns
            If CheckTag(iColumn, strTagFilter) = True Then
                If iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("N") Then
                    If iColumn.DefaultCellStyle.Format.Contains("0") Then
                        Cells.Add(New XMLExcelCell("", ExcelStringFormats.Decimal0))
                    Else
                        If iColumn.DefaultCellStyle.Format.Contains("1") Then
                            Cells.Add(New XMLExcelCell("", ExcelStringFormats.Decimal1))
                        Else
                            If iColumn.DefaultCellStyle.Format.Contains("2") Then
                                Cells.Add(New XMLExcelCell("", ExcelStringFormats.Decimal2))
                            Else
                                If iColumn.DefaultCellStyle.Format.Contains("3") Then
                                    Cells.Add(New XMLExcelCell("", ExcelStringFormats.Decimal3))
                                Else
                                    Cells.Add(New XMLExcelCell("", ExcelStringFormats.Decimal4))
                                End If
                            End If
                        End If
                    End If
                Else
                    If iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("C") Then
                        Cells.Add(New XMLExcelCell("", ExcelStringFormats.Currency2))
                    Else
                        If iColumn.DefaultCellStyle.Format.ToUpper.StartsWith("P") Then
                            If iColumn.DefaultCellStyle.Format.Contains("0") Then
                                Cells.Add(New XMLExcelCell("", ExcelStringFormats.Percent0))
                            Else
                                If iColumn.DefaultCellStyle.Format.Contains("2") Then
                                    Cells.Add(New XMLExcelCell("", ExcelStringFormats.Percent2))
                                Else
                                    Cells.Add(New XMLExcelCell("", ExcelStringFormats.Percent4))
                                End If
                            End If
                        Else
                            Cells.Add(New XMLExcelCell("", ExcelStringFormats.NoStyle)) '  .Normal10))
                        End If
                    End If
                End If
            End If
        Next

        For Each iRow As DataGridViewRow In dg.Rows

            'Try to fish out the last row which is just a row of 0's.
            If iRow.IsNewRow = False Then
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
                    If blnCurrentBold = True And iRow.Index = dg.CurrentRow.Index Then

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
                    If CheckTag(iColumn, strTagFilter) = True Then
                        If Not iRow.Cells(iDataGridColumn).Value Is System.DBNull.Value Then

                            '20130522 TimeSpan was causing an exception.
                            Try
                                Cells(iCellsColumn).strValue = iRow.Cells(iDataGridColumn).FormattedValue   '.Value
                            Catch ex As Exception
                                Dim ts As TimeSpan = iRow.Cells(iDataGridColumn).Value
                                Cells(iCellsColumn).strValue = ts.ToString()
                            End Try


                            '20120314 Set the cell style to bold if the cell is bold.
                            Try
                                If Not iRow.Cells(iDataGridColumn).Style Is Nothing Then
                                    If Not iRow.Cells(iDataGridColumn).Style.Font Is Nothing Then
                                        If iRow.Cells(iDataGridColumn).Style.Font.Bold = True Then
                                            Cells(iCellsColumn).blnBold = True
                                        End If
                                    End If
                                End If
                            Catch ex As Exception
                            End Try
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

                WriteStringToExcel(Cells, 1, False, iRowStyle)

            End If
        Next

        'Adjust ColumnCount so that it includes all columns written.
        AdjustColumnCount(iCellsColumn)

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

    Public Sub SetRow(ByVal iRow As Integer)
        RowCount = iRow
    End Sub
#End Region

#Region "WriteString"
    Public Sub SetVerticalPageBreak(ByVal iColumn As Integer)
        'oSheet.Columns(iColumn).Select()
        'oSheet.VPageBreaks.Add(xlApp.Sheets.Application.ActiveCell)
        ''oSheet.VPageBreaks(iBreak).Location = oSheet.Range(strCoordinate)
    End Sub

    Public Sub SetHorizontalPageBreak(ByVal iRow As Integer)
        'oSheet.Rows(iRow).Select()
        'oSheet.HPageBreaks.Add(xlApp.Sheets.Application.ActiveCell)
        ''oSheet.HPageBreaks(iBreak).Location = oSheet.Range(strCoordinate)
    End Sub

    Public Sub OpenMergedRow(ByVal iStyleName As ExcelStringFormats, ByVal iColumnHeight As Integer)
        Dim strStyle = strGetStyle(iStyleName)
        StoreXMLData("<Row ss:Height='" & iColumnHeight.ToString() & "' ss:StyleID='" & strStyle & "'>")
    End Sub

    Public Sub WriteStringToExcelAndMerge(ByVal strT As String, ByVal iStyleName As ExcelStringFormats, _
    ByVal strT2 As String, ByVal iStyleName2 As ExcelStringFormats, ByVal iMergeColumns As Integer)

        'Writes 2 strings with different style on same row and merges 2nd string with rows.
        OpenMergedRow(iStyleName, 14 + (14 * strT2.Length) / 120)
        Dim iCol As Integer = 1
        Dim strSplit() As String
        Dim s As String
        strSplit = strT.Split(New [Char]() {"#"c})
        Dim strStyle = strGetStyle(iStyleName)
        Dim strType = strGetType(iStyleName)
        For Each s In strSplit
            pWriteCell(s, strStyle, strType)
            iCol = iCol + 1
        Next s

        strSplit = strT2.Split(New [Char]() {"#"c})
        strStyle = strGetStyle(iStyleName2)
        strType = strGetType(iStyleName2)
        For Each s In strSplit
            Dim strValue As String
            '20081112 Filter @, < and >.
            strValue = FilterXML(s)
            Dim strVal As String = "<Cell ss:MergeAcross= '" & iMergeColumns.ToString() & "'"
            If strStyle <> "" Then
                strVal = strVal + " ss:StyleID='" & strStyle & "'>"
            Else
                strVal = strVal + ">"
            End If

            If strValue.Length <> 0 Then
                strVal = strVal & "<Data ss:Type='" & strType & "'>" & strValue & "</Data>"
            End If
            strVal = strVal + "<NamedCell ss:Name='Print_Area'/></Cell>"
            StoreXMLData(strVal)

            'pWriteCell(s, strStyle, strType)
            iCol = iCol + 1
        Next s

        CloseRow()
        AdjustColumnCount(iCol)
        AdjustRowCount(1)

        'Reset the row colour.
        SetRowColour(0)

    End Sub

    Public Sub SetRowColour(ByVal iIndex As Integer)

        'Call before Writing a string.
        iColourIndex = iIndex
    End Sub

    Public Sub WriteStringToExcel(ByVal strT As String, ByVal iStyleName As ExcelStringFormats, ByVal strT2 As String, ByVal iStyleName2 As ExcelStringFormats)

        'Writes 2 strings with different style on same row.
        OpenRow(iStyleName, -1)
        Dim iCol As Integer = 1
        Dim strSplit() As String
        Dim s As String
        strSplit = strT.Split(New [Char]() {"#"c})
        Dim strStyle = strGetStyle(iStyleName)
        Dim strType = strGetType(iStyleName)
        For Each s In strSplit
            pWriteCell(s, strStyle, strType)
            iCol = iCol + 1
        Next s
        strSplit = strT2.Split(New [Char]() {"#"c})
        strStyle = strGetStyle(iStyleName2)
        strType = strGetType(iStyleName2)
        For Each s In strSplit
            pWriteCell(s, strStyle, strType)
            iCol = iCol + 1
        Next s

        CloseRow()
        AdjustColumnCount(iCol)
        AdjustRowCount(1)

        'Reset the row colour.
        SetRowColour(0)

    End Sub

    Public Sub WriteStringToExcel(ByVal strT As String, ByVal iStyleName As ExcelStringFormats, ByVal iColumnHeight As Integer)
        WriteStringToExcel(strT, iStyleName, 1, False, iColumnHeight)
    End Sub

    Public Sub WriteStringToExcel(ByVal strT As String, ByVal iStyleName As ExcelStringFormats)
        WriteStringToExcel(strT, iStyleName, 1, False)
    End Sub

    Public Function WriteStringToExcel(ByVal strT As String, _
    ByVal iStyleName As ExcelStringFormats, _
    ByVal iStartColumn As Integer, _
    ByVal blnRightAlign As Boolean) As Integer
        Return WriteStringToExcel(strT, iStyleName, 1, False, -1)
    End Function

    Public Function WriteStringToExcel(ByVal strT As String, _
    ByVal iStyleName As ExcelStringFormats, _
    ByVal iStartColumn As Integer, _
    ByVal blnRightAlign As Boolean, ByVal iColumnHeight As Integer) As Integer

        ''Write the string to Excel using a separate column for each field in the string.
        ''Field separator is #.
        'All the fields are written using the same style.
        Dim strSplit() As String
        Dim s As String
        Dim iCol As Integer

        strSplit = strT.Split(New [Char]() {"#"c})
        iCol = iStartColumn
        Dim strStyle = strGetStyle(iStyleName)
        Dim strType = strGetType(iStyleName)

        OpenRow(iStyleName, iColumnHeight)
        For Each s In strSplit
            pWriteCell(s, strStyle, strType)
            iCol = iCol + 1
        Next s
        CloseRow()
        AdjustColumnCount(iCol)
        AdjustRowCount(1)

        'Reset the row colour.
        SetRowColour(0)
        Return iCol
    End Function

    Public Sub OpenRow(ByVal iStyleName As ExcelStringFormats, ByVal iColumnHeight As Integer)
        Dim strStyle = strGetStyle(iStyleName)
        If iColumnHeight = -1 Then
            'StoreXMLData("<Row ss:StyleID='" & strStyle & "'>")
            '20110423 RPB. Add AutofitHeight=1 to allow the rows to adjust to word wrapped text. 
            'See also modification of 20110423 to the Style.
            If blnWordWrap = True Then
                StoreXMLData("<Row ss:AutoFitHeight='1' ss:StyleID='" & strStyle & "'>")
            Else
                StoreXMLData("<Row ss:StyleID='" & strStyle & "'>")
            End If
        Else
            StoreXMLData("<Row ss:AutoFitHeight='0' ss:Height='" & iColumnHeight.ToString() & "' ss:StyleID='" & strStyle & "'>")
        End If

    End Sub

    Public Sub CloseRow()
        StoreXMLData("</Row>")
        'stW.WriteLine("</Row>")
    End Sub

    Public Function WriteStringToExcel(ByVal Cells As List(Of XMLExcelCell), _
        ByVal iStartColumn As Integer, _
        ByVal blnRightAlign As Boolean, ByVal iStyle As ExcelStringFormats) As Integer

        'Write the string to Excel using a separate column for each field in the string.
        Dim iCol As Integer
        iCol = iStartColumn
        Dim c As XMLExcelCell
        Dim strStyle As String
        Dim strType As String

        OpenRow(iStyle, -1)
        For Each c In Cells

            'Mod RPB June 2008. Dg could contain nulls. So replace with "" to avoid gaps.
            If c.strValue Is Nothing Then
                c.strValue = ""
            End If

            '20120314And c.iStyle = ExcelStringFormats.Normal10 
            If c.blnBold = True And c.iStyle = ExcelStringFormats.NoStyle Then
                strStyle = strGetStyle(ExcelStringFormats.Bold10)
            Else
                strStyle = strGetStyle(c.iStyle)
            End If
            strType = strGetType(c.iStyle)
            pWriteCell(c.strValue, strStyle, strType)
            iCol = iCol + 1
        Next c
        CloseRow()
        AdjustColumnCount(iCol)
        AdjustRowCount(1)
        Return iCol
    End Function

    Public Function WriteStringToExcel(ByVal Cells As List(Of XMLExcelCell), _
        ByVal iStartColumn As Integer, _
        ByVal blnRightAlign As Boolean) As Integer

        'Write the string to Excel using the style of the first cell as the default row style.
        'However it will not be used anyway because each cell has its own style.
        Dim iCol As Integer = -1
        For Each c As XMLExcelCell In Cells
            iCol = WriteStringToExcel(Cells, iStartColumn, blnRightAlign, c.iStyle)
            Exit For
        Next
        Return iCol

    End Function

    Public Sub WriteCell(ByVal strValue As String, ByVal iStyleName As ExcelStringFormats, ByVal strType As String)

        'Public override.
        Dim strStyle As String
        strStyle = strGetStyle(iStyleName)
        pWriteCell(strValue, strStyle, strType)
    End Sub

    Private Function FilterXML(ByVal strS As String) As String
        'XML strings should not contain & < or >  so filter out and replace with the XML escape characters.
        Dim strRet As String
        strRet = strS.Replace("&", "&amp;")
        strRet = strRet.Replace("<", "&lt;")
        strRet = strRet.Replace(">", "&gt;")
        Return strRet
    End Function

    Dim strDefaultType As String = "String"
    Private Sub pWriteCell(ByVal strValue As String, ByVal strStyle As String, ByVal strType As String)

        If strType = "DateTime" Then
            If strValue.Length = 0 Then
                strType = strDefaultType
            Else
                Dim dt As DateTime
                dt = strValue
                strValue = dt.Year & "-" & dt.Month.ToString("0#") & "-" & dt.Day.ToString("0#") & "T00:00:00.000"
            End If
        End If

        'Check whether we really do have a number. If not change to String.
        '20110602 RPB Modified pWriteCell. Check whether a number is really a number before replacing , by . because System.Convert.ToDouble is sensitive to regional setting.
        If strType = "Number" Then
            Try
                '20110504
                If strValue.Length > 0 Then
                    Dim dD As Double = System.Convert.ToDouble(strValue)
                End If
            Catch ex As Exception
                strType = strDefaultType
                strStyle = "NormalRightAligned"
            End Try
        End If

        'See comment on regional settings above.
        '20081105 Added this after testing on US Excel.
        If strType = "Number" And strSeparator = "," Then
            'originalCulture = System.Threading.Thread.CurrentThread.CurrentCulture
            strValue = strValue.Replace(".", "#")
            strValue = strValue.Replace(",", ".")
            strValue = strValue.Replace("#", ",")
        End If

        '20081112 Filter @, < and >.
        strValue = FilterXML(strValue)
        Dim strVal As String
        If strStyle <> "" Then
            strVal = "<Cell ss:StyleID='" & strStyle & "'>"
        Else
            strVal = "<Cell>"
        End If

        '20111118 Users can alter value for True and False.
        If strValue = "True" Then strValue = _strTrue
        If strValue = "False" Then strValue = _strFalse

        If strValue.Length <> 0 Then
            strVal = strVal & "<Data ss:Type='" & strType & "'>" & strValue & "</Data>"
        End If
        strVal = strVal + "<NamedCell ss:Name='Print_Area'/></Cell>"
        StoreXMLData(strVal)
        'stW.WriteLine(strVal)
    End Sub

#End Region

#Region "Layout"
    Public Sub SetAutofit(ByVal ColumnWidth As List(Of Integer))
        For Each i As Integer In ColumnWidth
            StoreXMLData("<Column ss:Width='" & i.ToString & "'/>")
            'StoreXMLData("<Column ss:Width='" & i.ToString & "' ss:StyleID='s999'/>")
            'StoreXMLData("<Column ss:Width='" & i.ToString & "' ss:WrapText='1' />")
        Next
    End Sub

    Public Sub SetAutofit(ByVal ColumnWidth As List(Of Integer), ByVal ColumnHidden As List(Of Boolean) _
        , ByVal ColumnNumber As Integer)

        'Write column hiding if necessary.
        For i As Integer = 0 To (ColumnNumber - 1)
            If ColumnHidden(i) = True Then
                StoreXMLData("<Column ss:Hidden='1' ss:AutoFitWidth='0' ss:Width='" & ColumnWidth(i).ToString & "' ss:Span='0'/>")
            Else
                StoreXMLData("<Column ss:Width='" & ColumnWidth(i).ToString & "'/>")
            End If
        Next
        'For Each i As Integer In ColumnWidth

        '    '<Column ss:Hidden="1" ss:AutoFitWidth="0" ss:Width="75" ss:Span="2"/>

        '    StoreXMLData("<Column ss:Width='" & i.ToString & "'/>")
        '    'stW.WriteLine("<Column ss:Width='" & i.ToString & "'/>")
        'Next
    End Sub

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

    Public Sub SetFitTo1Page()
        blnFitToPage = True
        'oSheet.PageSetup.FitToPagesWide = 1
        'oSheet.PageSetup.FitToPagesTall = 1
    End Sub

    Private Function PageSetup(ByVal strHeaderText As String, ByVal blnLandscape As Boolean, ByVal blnShowHeader As Boolean, _
    ByVal dLeftMargin As Double, ByVal dRightMargin As Double, ByVal dTopMargin As Double, ByVal dBottomMargin As Double, _
    ByVal dHeaderMargin As Double, _
    ByVal dFooterMargin As Double, ByVal iPaperSize As Integer, ByVal blnFitToPage As Boolean, _
    ByVal blnRepeatTopRow As Boolean, ByVal iPrintQuality As Integer, ByVal iZoom As Integer) As Boolean

        Dim blnRet As Boolean
        ''Adjust the Excel page setup for landscape, displaying the HeaderText in the header of the print out. 
        'Try
        blnRet = True

        '    With oSheet.PageSetup
        '        If blnRepeatTopRow = True Then
        '            .PrintTitleRows = "$1:$1"
        '        Else
        '            .PrintTitleRows = ""
        '        End If

        '        .PrintTitleColumns = ""
        '        .LeftMargin = xlApp.Sheets.Application.InchesToPoints(dLeftMargin)
        '        .RightMargin = xlApp.Sheets.Application.InchesToPoints(dRightMargin)
        '        .TopMargin = xlApp.Sheets.Application.InchesToPoints(dTopMargin)
        '        .BottomMargin = xlApp.Sheets.Application.InchesToPoints(dBottomMargin)
        '        .HeaderMargin = xlApp.Sheets.Application.InchesToPoints(dHeaderMargin)
        '        .FooterMargin = xlApp.Sheets.Application.InchesToPoints(dFooterMargin)
        '        If blnShowHeader = True Then
        '            .LeftHeader = "Using data from " & strHeaderText & "."
        '            .CenterHeader = "Printed on &D &T. "
        '            .RightHeader = "Page &P of &N"
        '            .LeftFooter = ""
        '            .CenterFooter = ""
        '            .RightFooter = "File created by " & System.Windows.Forms.Application.ProductName & "."
        '        Else
        '            .LeftHeader = ""
        '            .CenterHeader = ""
        '            .RightHeader = ""
        '            .LeftFooter = ""
        '            .CenterFooter = ""
        '            .RightFooter = ""
        '        End If
        '        .PrintHeadings = False
        '        .PrintGridlines = False
        '        .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
        '        .CenterHorizontally = False
        '        .CenterVertically = False
        '        If blnLandscape = True Then
        '            .Orientation = Excel.XlPageOrientation.xlLandscape
        '        Else
        '            .Orientation = Excel.XlPageOrientation.xlPortrait
        '        End If
        '        .Draft = False
        '        If iPaperSize <> 0 Then
        '            .PaperSize = iPaperSize     'Excel.XlPaperSize.xlPaperA4
        '        End If
        '        If iPrintQuality <> 0 Then
        '            .PrintQuality = iPrintQuality
        '        End If

        '        '.FirstPageNumber = xlAutomatic
        '        .Order = Excel.XlOrder.xlDownThenOver
        '        .BlackAndWhite = False
        '        If blnFitToPage = True Then
        '            .Zoom = 70
        '            .Zoom = False   'activate the Fit parameters.
        '            .FitToPagesWide = 1
        '            .FitToPagesTall = False
        '        End If
        '        If iZoom = 0 Then
        '            .Zoom = False
        '        Else
        '            .Zoom = iZoom
        '        End If
        '        '            .PrintArea = strPrtArea '"R1C1:R26C184"
        '    End With
        'Catch ex As Exception
        '    MsgBox("Could not set up the page. Check whether printer was selected. " & ex.Message, MsgBoxStyle.OkOnly)
        '    blnRet = False
        'Finally
        '    ColumnCount = 1
        '    RowCount = 1
        'End Try
        Return blnRet
    End Function

#End Region

#Region "InsertPicture"

    Public Sub InsertPicture(ByVal strSheetName As String, ByVal strImageName As String, ByVal strAtCell As String)
        If Not wb Is Nothing Then
            For Each sheet As Worksheet In wb.Sheets
                If sheet.Name = strSheetName Then
                    Dim p As Object = sheet.Pictures.Insert(strImageName)
                    p.left = sheet.Range(strAtCell).Left
                    p.top = sheet.Range(strAtCell).Top

                    p = Nothing
                End If
            Next
        End If
    End Sub

    Public Sub InsertPicture(ByVal strSheetName As String, ByVal strImageName As String, ByVal strAtCell As String, width As Object, height As Object)
        If Not wb Is Nothing Then
            For Each sheet As Worksheet In wb.Sheets
                If sheet.Name = strSheetName Then
                    Dim p As Object = sheet.Pictures.Insert(strImageName)
                    p.left = sheet.Range(strAtCell).Left
                    p.top = sheet.Range(strAtCell).Top
                    p.height = height
                    p.width = width
                    p = Nothing
                End If
            Next
        End If
    End Sub

    Public Sub InsertPicture(ByVal strImageName As String)
        ''xlApp.ActiveSheet.Pictures.Insert(strImageName) '"T:\Everyone\RAP\RAP2\Output\JPGs\image002.jpg"). _
        'xlApp.Sheets.Application.ActiveSheet.Pictures.Insert(strImageName) '"T:\Everyone\RAP\RAP2\Output\JPGs\image002.jpg"). _
    End Sub
#End Region

#Region "Formula"

    Public Sub MakeASum(ByVal iAtRow As Integer, ByVal iAtColumn As Integer, ByVal iStyleName As ExcelStringFormats, _
        ByVal iFromRow As Integer, ByVal iToRow As Integer, ByVal dDivideBy As Double)

        'Write a sum or sum and divide by.
        Dim strStyle As String
        strStyle = strGetStyle(iStyleName)
        Dim strFormula As String
        strFormula = "<Cell ss:StyleID='" & strStyle & "' ss:Formula='" & _
            "=Sum(R" & iFromRow & "C" & iAtColumn & ": R" & iToRow & "C" & iAtColumn & ")"
        If dDivideBy <> 0 And dDivideBy <> 1 Then
            strFormula = strFormula & " / " + dDivideBy.ToString()
        End If
        strFormula = strFormula & "'><NamedCell ss:Name='Print_Area'/></Cell>"
        StoreXMLData(strFormula)
        'stW.WriteLine(strFormula)

    End Sub

    Public Sub CreateFormula(ByVal iAtRow As Integer, ByVal iAtColumn As Integer, ByVal iStyleName As ExcelStringFormats, _
    ByVal strFormula As String)

        'xlApp.Sheets.Application.Cells(iAtRow, iAtColumn).Formula = strFormula
        'Write a sum or sum and divide by.
        Dim strStyle As String
        strStyle = strGetStyle(iStyleName)
        Dim strValue As String
        strValue = "<Cell ss:StyleID='" & strStyle & "' ss:Formula='" & strFormula & "'><NamedCell ss:Name='Print_Area'/></Cell>"
        StoreXMLData(strValue)
        'stW.WriteLine(strValue)
    End Sub

#End Region

    Private Function strPrecision(ByVal strFormat As String) As String

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
    Friend ExcelFormat

End Class
