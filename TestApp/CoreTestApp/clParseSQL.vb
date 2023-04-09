'------------------------------------------------
'Name: clParseSQL.vb.
'Function: 
'Created Jan 2016.
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
Imports Utilities
Imports System.IO
Imports System.Collections
'Imports Microsoft.Data.Schema.ScriptDom.Sql
'Imports Microsoft.SqlServer.Management.Smo
Public Class clParseSQL
    Friend WithEvents MetaData As CoreTestApp.MetaData
    Friend WithEvents typesTableAdapter As CoreTestApp.MetaDataTableAdapters.typesTableAdapter
    Friend WithEvents m_sql_charactersTableAdapter As CoreTestApp.MetaDataTableAdapters.m_sql_charactersTableAdapter
    Friend WithEvents m_keywordsTableAdapter As CoreTestApp.MetaDataTableAdapters.m_keywordsTableAdapter
    Dim tbl_types As CoreTestApp.MetaData.typesDataTable
    Dim tbl_m_sql_characters As CoreTestApp.MetaData.m_sql_charactersDataTable
    Dim tbl_m_keywords As CoreTestApp.MetaData.m_keywordsDataTable
    Public Sub New()
        MetaData = New CoreTestApp.MetaData()
        Me.typesTableAdapter = New CoreTestApp.MetaDataTableAdapters.typesTableAdapter()
        Me.m_sql_charactersTableAdapter = New CoreTestApp.MetaDataTableAdapters.m_sql_charactersTableAdapter()
        Me.m_keywordsTableAdapter = New CoreTestApp.MetaDataTableAdapters.m_keywordsTableAdapter()
        Me.typesTableAdapter.ClearBeforeFill = True
        m_sql_charactersTableAdapter.ClearBeforeFill = True
        Me.m_keywordsTableAdapter.ClearBeforeFill = True

        typesTableAdapter.Connection.ConnectionString = ConnectionString.ConnectionString
        m_sql_charactersTableAdapter.Connection.ConnectionString = ConnectionString.ConnectionString
        m_keywordsTableAdapter.Connection.ConnectionString = ConnectionString.ConnectionString

        tbl_types = typesTableAdapter.GetData()
        tbl_m_sql_characters = m_sql_charactersTableAdapter.GetData()
        tbl_m_keywords = m_keywordsTableAdapter.GetData()
    End Sub

    Private Function IsType(strT As String) As Boolean
        For Each tbl_typeRow As CoreTestApp.MetaData.typesRow In tbl_types
            If tbl_typeRow.name.ToUpper() = strT.ToUpper() Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function IsKeyword(strT As String) As Boolean
        For Each tbl_m_keywordsRow As CoreTestApp.MetaData.m_keywordsRow In tbl_m_keywords
            If tbl_m_keywordsRow.keyword.ToUpper() = strT.ToUpper() Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Sub ParseRTB(rtb As RichTextBox)
        Dim StrChecked As String = ""
        rtb.SelectAll()
        rtb.SelectionColor = Color.Green '.DarkGreen 'Color.Black

        Dim iLines As Integer = 0
        Dim iIndex As Integer
        Dim strT As String = rtb.Text.ToUpper()

        'look for ( and then highlight the word before if no ..
        'Or rtb.Text.Substring(iIndex + tbl_typesRow.name.Length, 1) = "("
        Dim strTT As String = ""
        iIndex = 0
        While iIndex > -1
            iIndex = rtb.Text.IndexOf("(", iIndex)
            strTT = ""
            If iIndex > 0 Then
                For k As Integer = iIndex - 1 To 0 Step -1
                    strTT = strTT + rtb.Text(k).ToString()
                    If Not Char.IsLetter(rtb.Text(k)) And Not rtb.Text(k) = "_" Then
                        If k < (iIndex - 1) And Not rtb.Text(k) = "." Then
                            rtb.Select(k, iIndex - k)
                            rtb.SelectionColor = Color.Fuchsia
                        End If
                        Exit For
                    End If
                Next
                iIndex = rtb.Text.IndexOf("(", iIndex + 1)
            End If
        End While

        'sql keywords
        For Each tbl_m_keywordsRow As CoreTestApp.MetaData.m_keywordsRow In tbl_m_keywords
            iIndex = 0
            While iIndex > -1
                iIndex = strT.IndexOf(tbl_m_keywordsRow.keyword.ToUpper(), iIndex)
                If iIndex > -1 Then
                    'check whether this is a complete word.
                    If String.IsNullOrWhiteSpace(rtb.Text.Substring(iIndex + tbl_m_keywordsRow.keyword.Length, 1)) Then
                        If iIndex - 1 < 0 Then
                            rtb.Select(iIndex, tbl_m_keywordsRow.keyword.Length)
                            rtb.SelectionColor = Color.Blue
                        Else
                            '(select
                            If String.IsNullOrWhiteSpace(rtb.Text.Substring(iIndex - 1, 1)) Or rtb.Text.Substring(iIndex - 1, 1) = "(" Then
                                rtb.Select(iIndex, tbl_m_keywordsRow.keyword.Length)
                                rtb.SelectionColor = Color.Blue
                            End If
                        End If
                    End If
                    iIndex += 1
                End If
            End While
        Next

        'sql types
        For Each tbl_typesRow As CoreTestApp.MetaData.typesRow In tbl_types
            iIndex = 0
            While iIndex > -1
                iIndex = strT.IndexOf(tbl_typesRow.name.ToUpper(), iIndex)
                If iIndex > -1 Then
                    'check whether this is a complete word.
                    If iIndex + tbl_typesRow.name.Length = rtb.Text.Length Then
                        rtb.Select(iIndex, tbl_typesRow.name.Length)
                        rtb.SelectionColor = Color.Blue
                    Else
                        If (String.IsNullOrWhiteSpace(rtb.Text.Substring(iIndex + tbl_typesRow.name.Length, 1)) Or
                            rtb.Text.Substring(iIndex + tbl_typesRow.name.Length, 1) = "(" Or
                             rtb.Text.Substring(iIndex + tbl_typesRow.name.Length, 1) = ")") And
                         String.IsNullOrWhiteSpace(rtb.Text.Substring(iIndex - 1, 1)) Then
                            rtb.Select(iIndex, tbl_typesRow.name.Length)
                            rtb.SelectionColor = Color.Blue
                        End If
                    End If
                    iIndex += 1
                End If
            End While
        Next

        'sql operators
        Dim strLastLetter As String = "1"
        For Each tbl_m_sql_charactersRow As CoreTestApp.MetaData.m_sql_charactersRow In tbl_m_sql_characters
            iIndex = 0
            While iIndex > -1
                iIndex = strT.IndexOf(tbl_m_sql_charactersRow.characters.ToUpper(), iIndex)
                If iIndex > -1 Then

                    'check whether this is a complete word.
                    If iIndex + tbl_m_sql_charactersRow.characters.Length = rtb.Text.Length Then
                        rtb.Select(iIndex, tbl_m_sql_charactersRow.characters.Length)
                        rtb.SelectionColor = Color.Black
                    Else
                        If String.IsNullOrWhiteSpace(rtb.Text.Substring(iIndex + tbl_m_sql_charactersRow.characters.Length, 1)) Or
                            tbl_m_sql_charactersRow.characters.Length = 1 Then

                            'check previous character is not a letter in which case this word is part of another.
                            If iIndex = 0 Then
                                rtb.Select(iIndex, tbl_m_sql_charactersRow.characters.Length)
                                rtb.SelectionColor = Color.Black
                            Else
                                If rtb.Text.Substring(iIndex - 1, 1) = " " Or rtb.Text.Substring(iIndex - 1, 1) = vbCrLf Or rtb.Text.Substring(iIndex - 1, 1) = vbLf Or rtb.Text.Substring(iIndex - 1, 1) = vbTab Then
                                    rtb.Select(iIndex, tbl_m_sql_charactersRow.characters.Length)
                                    rtb.SelectionColor = Color.Black
                                End If
                            End If
                        End If
                    End If
                    iIndex += 1
                End If
            End While
        Next

        'look for number and then highlight the following word as well.
        'strLastLetter = "1"
        'For k As Integer = 0 To rtb.Text.Length - 1
        '    If Char.IsNumber(rtb.Text(k)) Then
        '        If Not Char.IsLetter(strLastLetter(0)) And strLastLetter <> "_" Then
        '            rtb.Select(k, 1)
        '            rtb.SelectionColor = Color.Black
        '        End If
        '    Else
        '        strLastLetter = (rtb.Text(k))
        '    End If
        'Next


        'look for @@ and then highlight the following word as well.
        iIndex = 0
        While iIndex > -1
            iIndex = rtb.Text.IndexOf("@@", iIndex)
            If iIndex > -1 Then
                For k As Integer = iIndex To rtb.Text.Length
                    If String.IsNullOrWhiteSpace(rtb.Text(k)) Or rtb.Text(k) = vbCrLf Or rtb.Text(k) = vbLf Or rtb.Text(k) = vbTab Then
                        rtb.Select(iIndex, k - iIndex)
                        rtb.SelectionColor = Color.Fuchsia
                        Exit For
                    End If
                Next
                iIndex = rtb.Text.IndexOf("@@", iIndex + 1)
            End If
        End While

        '/**/
        'iIndex = 0
        'While iIndex > -1
        '    iIndex = rtb.Text.IndexOf("/*", iIndex)
        '    If iIndex > -1 Then
        '        For k As Integer = iIndex To rtb.Text.Length
        '            If rtb.Text.Substring(k, 2) = "*/" Then
        '                rtb.Select(iIndex, k + 2 - iIndex)
        '                rtb.SelectionColor = Color.ForestGreen
        '                Exit For
        '            End If
        '        Next
        '        iIndex += 1
        '    End If
        'End While

        'set the comments and "SingleLineComment" .
        'iLines = 0
        'For Each Str As String In rtb.Lines
        '    'If Str.StartsWith("--") Then
        '    iIndex = Str.IndexOf("--", 0)
        '    If iIndex > -1 Then
        '        rtb.Select(StrChecked.Length + iIndex + iLines, Str.Length - iIndex)
        '        rtb.SelectionColor = Color.ForestGreen
        '    End If
        '    StrChecked += Str
        '    iLines += 1
        'Next

        iIndex = 0
        strLastLetter = "1"
        For k As Integer = 0 To rtb.Text.Length - 1

            If k < rtb.Text.Length - 3 Then
                If rtb.Text.Substring(k, 2) = "/*" Then
                    iIndex = 0
                    For j As Integer = k + 1 To rtb.Text.Length - 1
                        If rtb.Text.Substring(j, 2) = "*/" Then
                            rtb.Select(k, j + 2 - k)
                            rtb.SelectionColor = Color.ForestGreen
                            iIndex = j + 1
                            strLastLetter = "1"
                            Exit For
                        End If
                        iIndex = j + 1
                    Next
                    k = iIndex
                End If
            End If
            If k < rtb.Text.Length - 3 Then
                If rtb.Text.Substring(k, 2) = "--" Then
                    iIndex = 0
                    For j As Integer = k + 1 To rtb.Text.Length - 1
                        If rtb.Text(j) = vbCrLf Or rtb.Text(j) = vbLf Or rtb.Text(j) = vbTab Then
                            rtb.Select(k, j - k)
                            rtb.SelectionColor = Color.ForestGreen
                            iIndex = j
                            strLastLetter = "1"
                            Exit For
                        End If
                        iIndex = j
                    Next
                    k = iIndex
                End If
            End If

            If rtb.Text.Substring(k, 1) = "'" Then
                iIndex = 0
                For j As Integer = k + 1 To rtb.Text.Length - 1
                    If rtb.Text.Substring(j, 1) = "'" Then
                        rtb.Select(k, j - k + 1)
                        rtb.SelectionColor = Color.Red
                        iIndex = j
                        strLastLetter = "1"
                        Exit For
                    End If
                    iIndex = j
                Next
                k = iIndex
            End If


            If Char.IsNumber(rtb.Text(k)) Then
                If Not Char.IsLetter(strLastLetter(0)) And strLastLetter <> "_" Then
                    rtb.Select(k, 1)
                    rtb.SelectionColor = Color.Black
                End If
            Else
                strLastLetter = (rtb.Text(k))
            End If

            'If k >= (rtb.Text.Length - 1) Then
            '    Exit For
            'End If
            'If k >= 1367 Then
            '    rtb.Select(k - 2, 1)
            '    rtb.SelectionColor = Color.Red
            '    '' Exit For
            'End If
        Next

        'iIndex = 0
        'While iIndex > -1
        '    iIndex = rtb.Text.IndexOf("/*", iIndex)
        '    If iIndex > -1 Then
        '        For k As Integer = iIndex To rtb.Text.Length
        '            If rtb.Text.Substring(k, 2) = "*/" Then
        '                rtb.Select(iIndex, k + 2 - iIndex)
        '                rtb.SelectionColor = Color.ForestGreen
        '                Exit For
        '            End If
        '        Next
        '        iIndex += 1
        '    End If
        'End While

        'Set the "AsciiStringLiteral"
        'Dim StrTextToGo As String = rtb.Text
        'Dim StrTextDone As String = ""
        'iIndex = StrTextToGo.IndexOf("'", 0)
        'While iIndex > -1
        '    StrTextDone = StrTextDone + StrTextToGo.Substring(0, iIndex + 1)
        '    StrTextToGo = StrTextToGo.Substring(iIndex + 1)
        '    iIndex = StrTextToGo.IndexOf("'", 0)
        '    If iIndex > -1 Then
        '        rtb.Select(StrTextDone.Length - 1, iIndex + 2)
        '        rtb.SelectionColor = Color.Red
        '        StrTextDone = StrTextDone + StrTextToGo.Substring(0, iIndex + 1)
        '        StrTextToGo = StrTextToGo.Substring(iIndex + 1)
        '        iIndex = StrTextToGo.IndexOf("'", 0)
        '    End If
        'End While


        'Use of the parser here is slow. The Types also need to be recognised separately as they are type Identifier.
        'Dim returnValue As ParseResult
        'Dim parseOptions = New ParseOptions
        ' Dim returnValue As IDictionary(Of String, Object)
        'returnValue = Parser.Parse(rtb.Text)
        'Dim parser As TSql100Parser = New TSql100Parser(True)
        'Dim rd As TextReader = New StringReader(rtb.Text)
        'Dim Errors = New List(Of Microsoft.Data.Schema.ScriptDom.ParseError)
        'Dim fragments As TSqlFragment = parser.Parse(rd, Errors)
        'Dim strTokens As String = ""
        'For i As Integer = fragments.FirstTokenIndex To fragments.LastTokenIndex
        '    If fragments.ScriptTokenStream(i).Text <> Nothing Then
        '        'If fragments.ScriptTokenStream(i).TokenType.ToString() <> "WhiteSpace" And _
        '        '    parser.ValidateIdentifier(fragments.ScriptTokenStream(i).Text) = False Then

        '        '    strTokens += fragments.ScriptTokenStream(i).Line.ToString()
        '        '    strTokens += "."
        '        '    strTokens += fragments.ScriptTokenStream(i).Offset.ToString()
        '        '    strTokens += " "
        '        '    strTokens += fragments.ScriptTokenStream(i).TokenType.ToString()
        '        '    strTokens += " "
        '        '    strTokens += fragments.ScriptTokenStream(i).Text
        '        '    strTokens += " "
        '        '    'strTokens += fragments.ScriptTokenStream(i).GetType().ToString()
        '        '    'strTokens += Environment.NewLine
        '        '    'strTokens += "               "

        '        'End If

        '        If fragments.ScriptTokenStream(i).TokenType.ToString() = "SingleLineComment" Then
        '            rtb.Select(fragments.ScriptTokenStream(i).Offset, fragments.ScriptTokenStream(i).Text.Length)
        '            rtb.SelectionColor = Color.ForestGreen
        '        Else
        '            If fragments.ScriptTokenStream(i).TokenType.ToString() = "AsciiStringLiteral" Then
        '                rtb.Select(fragments.ScriptTokenStream(i).Offset, fragments.ScriptTokenStream(i).Text.Length)
        '                rtb.SelectionColor = Color.Red
        '            Else
        '                If fragments.ScriptTokenStream(i).TokenType.ToString() = "Null" Or _
        '                    fragments.ScriptTokenStream(i).TokenType.ToString() = "Is" Or _
        '                    fragments.ScriptTokenStream(i).TokenType.ToString() = "Exists" Or _
        '                    fragments.ScriptTokenStream(i).TokenType.ToString() = "Star" Or _
        '                    fragments.ScriptTokenStream(i).TokenType.ToString() = "Comma" Or _
        '                    fragments.ScriptTokenStream(i).TokenType.ToString() = "And" Or _
        '                    fragments.ScriptTokenStream(i).TokenType.ToString() = "Or" Or _
        '                    fragments.ScriptTokenStream(i).TokenType.ToString() = "EqualsSign" Or _
        '                    fragments.ScriptTokenStream(i).TokenType.ToString() = "Integer" Or _
        '                    fragments.ScriptTokenStream(i).TokenType.ToString() = "RightParenthesis" Or _
        '                    fragments.ScriptTokenStream(i).TokenType.ToString() = "LeftParenthesis" Then
        '                    rtb.Select(fragments.ScriptTokenStream(i).Offset, fragments.ScriptTokenStream(i).Text.Length)
        '                    rtb.SelectionColor = Color.Black
        '                Else
        '                    'fragments.ScriptTokenStream(i).TokenType.ToString() <> "SingleLineComment" And _
        '                    '    fragments.ScriptTokenStream(i).TokenType.ToString() <> "Variable"
        '                    If parser.ValidateIdentifier(fragments.ScriptTokenStream(i).Text) = False And _
        '                        fragments.ScriptTokenStream(i).TokenType.ToString() <> "SingleLineComment" And _
        '                        fragments.ScriptTokenStream(i).TokenType.ToString() <> "Variable" And _
        '                        fragments.ScriptTokenStream(i).TokenType.ToString() <> "WhiteSpace" Then
        '                        rtb.Select(fragments.ScriptTokenStream(i).Offset, fragments.ScriptTokenStream(i).Text.Length)
        '                        rtb.SelectionColor = Color.Blue
        '                    End If
        '                End If
        '            End If
        '        End If
        '    End If
        'Next
        ''MsgBox(strTokens)
    End Sub


End Class
