'------------------------------------------------
'Name: Module for statics.vb.
'Function: Date conversion.
'Copyright Robin Baines 2010. All rights reserved.
'Created May 2010.
'Notes: 
'Modifications:
'------------------------------------------------
Imports System
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Security.Principal
Namespace dates
    Public Module ModuleDates
        Const strGlobalDateUndefined = "9999999" '
        Public Function strGetCJSDate() As String
            Dim strDate As String
            'strDate = DateTime.Today().Year.ToString("00") & DateTime.Today().Month.ToString("00") & DateTime.Today().Day.ToString("00")
            Dim dt As DateTime
            dt = System.DateTime.Today
            Dim iMonth = dt.Month
            'Dim iYear = dt.Year - 2000
            Dim iYear = dt.Year
            'strDate = "1" & Format(iYear, "00") & Format(iMonth, "00") & dt.Day.ToString("00")
            strDate = Format(iYear, "0000") & Format(iMonth, "00") & dt.Day.ToString("00")
            Return strDate
        End Function
        Public Function strGetCJSDate(ByVal dt As DateTime) As String
            Dim strDate As String
            'strDate = DateTime.Today().Year.ToString("00") & DateTime.Today().Month.ToString("00") & DateTime.Today().Day.ToString("00")
            Dim iMonth = dt.Month
            'Dim iYear = dt.Year - 2000
            Dim iYear = dt.Year
            'strDate = "1" & Format(iYear, "00") & Format(iMonth, "00") & dt.Day.ToString("00")
            strDate = Format(iYear, "0000") & Format(iMonth, "00") & dt.Day.ToString("00")
            Return strDate
        End Function
        Public Function strConvertFromCDateTime(ByVal strCDate As String, ByVal strCTime As String) As String

            'Converts "yyyymmdd hhmmss" to Short Date format.
            'So the Gregorian date has to be active. See Calendar setting in XMLTPI initialisation.
            '" " means the C Field was empty.
            Dim d As Date
            'Note that the following forces the date 'd' to the current short date format currently being used on
            'the computer. This is then used to check how the C date should be converted.
            Dim strDate As String

            d = #1/31/2003#
            strDate = Format(d, "Short Date")
            'First check whether the conversions will succeed.
            If InStr(strCDate, "/") = 0 Then
                If Len(Trim(strCDate)) <> 0 And strCDate <> strGlobalDateUndefined Then
                    If Left(strDate, 1) <> "3" Then  'mm/dd/yyyy'
                        strConvertFromCDateTime = _
                            Mid(strCDate, 5, 2) & "/" & Right(strCDate, 2) & _
                        "/" & Left(strCDate, 4)
                    Else
                        strConvertFromCDateTime = _
                             Right(strCDate, 2) & "/" & Mid(strCDate, 5, 2) & _
                        "/" & Left(strCDate, 4)
                    End If
                    If Len(Trim(strCTime)) <> 0 Then
                        strConvertFromCDateTime = strConvertFromCDateTime & _
                        " " & Left(strCTime, 2) & ":" & Mid(strCTime, 3, 2)
                    End If
                Else

                    'Mod RPB Aug 2005. This was " " and not "". This gave problems if the dates were not trimmed.
                    strConvertFromCDateTime = ""
                End If
            Else
                strConvertFromCDateTime = strCDate
                If Len(Trim(strCTime)) <> 0 Then
                    strConvertFromCDateTime = strConvertFromCDateTime & " " & strCTime
                End If
            End If

        End Function
        Public Function strConvertToTPDateTime(ByVal strAccessDate As String) As String

            '" " means the C Field was empty
            'Assumes C is sending US short date format and that TP server is also using this format.
            'mm/dd/yyyy = yyyymmdd
            If Len(Trim(strAccessDate)) Then
                strConvertToTPDateTime = Format(strAccessDate, "yyyymmdd")
            Else
                strConvertToTPDateTime = "" 'strGlobalCDateUndefined
            End If
        End Function
        Public Function strGetThisMonth() As String

            'This month returned in C format.
            Dim dt As DateTime
            dt = System.DateTime.Today
            Dim iMonth = dt.Month
            'Dim iYear = dt.Year - 2000
            Dim iYear = dt.Year
            'Return "1" & Format(iYear, "00") & Format(iMonth, "00")
            Return Format(iYear, "0000") & Format(iMonth, "00")
        End Function
        Public Function IsCDate(ByVal strDate As String) As Boolean

            'Return true if this is a C date yyyymmdd
            If Len(Trim(strDate)) = 8 And Left(Trim(strDate), 1) = "2" Then
                IsCDate = True
            Else
                IsCDate = False
            End If
        End Function
        Public Function strIncrementDate(ByVal strDate As String) As String

            'Increment the C date by 1 month.
            Dim strYear As String
            Dim strMonth As String
            Dim iYear As Integer
            Dim iMonth As Integer

            'strYear = Mid(strDate, 2, 2)
            strYear = Mid(strDate, 1, 4)
            'strMonth = Mid(strDate, 4, 2)
            strMonth = Mid(strDate, 5, 2)
            iYear = CInt(strYear)
            iMonth = CInt(strMonth)
            iMonth = iMonth + 1
            If iMonth > 12 Then
                iYear = iYear + 1

                iMonth = 1
            End If
            'strIncrementDate = "1" & Format(iYear, "00") & Format(iMonth, "00")
            strIncrementDate = Format(iYear, "0000") & Format(iMonth, "00")
        End Function

        '201206 created.
        Public Function strDecrementDate(ByVal strDate As String) As String

            'Increment the C date by 1 month.
            Dim strYear As String
            Dim strMonth As String
            Dim iYear As Integer
            Dim iMonth As Integer

            'strYear = Mid(strDate, 2, 2)
            strYear = Mid(strDate, 1, 4)
            'strMonth = Mid(strDate, 4, 2)
            strMonth = Mid(strDate, 5, 2)
            iYear = CInt(strYear)
            iMonth = CInt(strMonth)
            iMonth = iMonth - 1
            If iMonth < 1 Then
                iYear = iYear - 1

                iMonth = 12
            End If
            strDecrementDate = Format(iYear, "0000") & Format(iMonth, "00")
        End Function
        Public Function iDifferenceInDays(ByVal strCDate As String) As Integer
            Dim dt As DateTime
            dt = System.DateTime.Today
            Dim dt1 As DateTime = strConvertFromCDateTime(strCDate, "")
            Dim elapsedSpan As TimeSpan = New TimeSpan(dt1.Ticks - DateTime.Now.Date.Ticks)
            Return elapsedSpan.TotalDays
        End Function
    End Module
End Namespace