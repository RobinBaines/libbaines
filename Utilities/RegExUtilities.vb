'------------------------------------------------
'Name: Module for RegexUtilities.vb.
'Function: Check that a string is a valid email address.
'Copyright Robin Baines 2010. All rights reserved.
'Created May 2010.
'Notes: 
'Modifications:
'------------------------------------------------
Imports System.Globalization
Imports System.Text.RegularExpressions

Public Class RegexUtilities
    Dim invalid As Boolean = False

public Function IsValidEmail(strIn As String) As Boolean
        invalid = False
        If String.IsNullOrEmpty(strIn) Then Return False

        ' Use IdnMapping class to convert Unicode domain names.
        strIn = Regex.Replace(strIn, "(@)(.+)$", AddressOf Me.StrMatch)
        If invalid Then Return False

        ' Return true if strIn is in valid e-mail format.
        Return Regex.IsMatch(strIn, _
               "^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" + _
               "(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$",
               RegexOptions.IgnoreCase)
    End Function

    Private Function StrMatch(match As Match) As String
        '
        Dim idn As New IdnMapping()
        Dim StringToMatch As String = match.Groups(2).Value
        Try
            StringToMatch = idn.GetAscii(StringToMatch)
        Catch e As ArgumentException
            invalid = True
        End Try
        Return match.Groups(1).Value + StringToMatch
    End Function
End Class
