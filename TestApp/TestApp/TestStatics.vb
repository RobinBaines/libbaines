'------------------------------------------------
'Name: Module for TestStatics.vb.
'Function: 
'Created Jan 2011.
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
''Imports Microsoft.Data.Schema.ScriptDom.Sql

'Imports Microsoft.SqlServer.Management.Smo
Module TestStatics
    Enum ErrorLevels
        HighPriority = 10
        CatchError = 20
        LowPriority = 30
    End Enum

    'Public Sub ParseRTB(rtb As RichTextBox)
    '    Dim StrChecked As String = ""
    '    rtb.SelectAll()
    '    rtb.SelectionColor = Color.DarkGreen 'Color.Black

    '    'Dim iLines As Integer = 0
    '    'Dim iIndex As Integer
    '    'For Each Str As String In rtb.Lines
    '    '    'If Str.StartsWith("--") Then
    '    '    iIndex = Str.IndexOf("--", 0)
    '    '    If iIndex > -1 Then
    '    '        rtb.Select(StrChecked.Length + iIndex + iLines, Str.Length - iIndex)
    '    '        rtb.SelectionColor = Color.ForestGreen
    '    '    End If
    '    '    StrChecked += Str
    '    '    iLines += 1
    '    'Next
    '    'Dim StrTextToGo As String = rtb.Text
    '    'Dim StrTextDone As String = ""
    '    'iIndex = StrTextToGo.IndexOf("'", 0)
    '    'While iIndex > -1
    '    '    StrTextDone = StrTextDone + StrTextToGo.Substring(0, iIndex + 1)
    '    '    StrTextToGo = StrTextToGo.Substring(iIndex + 1)
    '    '    iIndex = StrTextToGo.IndexOf("'", 0)
    '    '    If iIndex > -1 Then
    '    '        rtb.Select(StrTextDone.Length - 1, iIndex + 2)
    '    '        rtb.SelectionColor = Color.Red
    '    '        StrTextDone = StrTextDone + StrTextToGo.Substring(0, iIndex + 1)
    '    '        StrTextToGo = StrTextToGo.Substring(iIndex + 1)
    '    '        iIndex = StrTextToGo.IndexOf("'", 0)
    '    '    End If
    '    'End While

    '    ' Dim returnValue As ParseResult
    '    'Dim parseOptions = New ParseOptions
    '    ' Dim returnValue As IDictionary(Of String, Object)
    '    'returnValue = Parser.Parse(rtb.Text)
    '    Dim parser As TSql100Parser = New TSql100Parser(True)
    '    Dim rd As TextReader = New StringReader(rtb.Text)
    '    Dim Errors = New List(Of Microsoft.Data.Schema.ScriptDom.ParseError)
    '    Dim fragments As TSqlFragment = parser.Parse(rd, Errors)
    '    Dim strTokens As String = ""
    '    For i As Integer = fragments.FirstTokenIndex To fragments.LastTokenIndex
    '        If fragments.ScriptTokenStream(i).Text <> Nothing Then
    '            'If fragments.ScriptTokenStream(i).TokenType.ToString() <> "WhiteSpace" And _
    '            '    parser.ValidateIdentifier(fragments.ScriptTokenStream(i).Text) = False Then

    '            '    strTokens += fragments.ScriptTokenStream(i).Line.ToString()
    '            '    strTokens += "."
    '            '    strTokens += fragments.ScriptTokenStream(i).Offset.ToString()
    '            '    strTokens += " "
    '            '    strTokens += fragments.ScriptTokenStream(i).TokenType.ToString()
    '            '    strTokens += " "
    '            '    strTokens += fragments.ScriptTokenStream(i).Text
    '            '    strTokens += " "
    '            '    'strTokens += fragments.ScriptTokenStream(i).GetType().ToString()
    '            '    'strTokens += Environment.NewLine
    '            '    'strTokens += "               "

    '            'End If

    '            If fragments.ScriptTokenStream(i).TokenType.ToString() = "SingleLineComment" Then
    '                rtb.Select(fragments.ScriptTokenStream(i).Offset, fragments.ScriptTokenStream(i).Text.Length)
    '                rtb.SelectionColor = Color.ForestGreen
    '            Else
    '                If fragments.ScriptTokenStream(i).TokenType.ToString() = "AsciiStringLiteral" Then
    '                    rtb.Select(fragments.ScriptTokenStream(i).Offset, fragments.ScriptTokenStream(i).Text.Length)
    '                    rtb.SelectionColor = Color.Red
    '                Else
    '                    If fragments.ScriptTokenStream(i).TokenType.ToString() = "Null" Or _
    '                        fragments.ScriptTokenStream(i).TokenType.ToString() = "Is" Or _
    '                        fragments.ScriptTokenStream(i).TokenType.ToString() = "Exists" Or _
    '                        fragments.ScriptTokenStream(i).TokenType.ToString() = "Star" Or _
    '                        fragments.ScriptTokenStream(i).TokenType.ToString() = "Comma" Or _
    '                        fragments.ScriptTokenStream(i).TokenType.ToString() = "And" Or _
    '                        fragments.ScriptTokenStream(i).TokenType.ToString() = "Or" Or _
    '                        fragments.ScriptTokenStream(i).TokenType.ToString() = "EqualsSign" Or _
    '                        fragments.ScriptTokenStream(i).TokenType.ToString() = "Integer" Or _
    '                        fragments.ScriptTokenStream(i).TokenType.ToString() = "RightParenthesis" Or _
    '                        fragments.ScriptTokenStream(i).TokenType.ToString() = "LeftParenthesis" Then
    '                        rtb.Select(fragments.ScriptTokenStream(i).Offset, fragments.ScriptTokenStream(i).Text.Length)
    '                        rtb.SelectionColor = Color.Black
    '                    Else
    '                        'fragments.ScriptTokenStream(i).TokenType.ToString() <> "SingleLineComment" And _
    '                        '    fragments.ScriptTokenStream(i).TokenType.ToString() <> "Variable"
    '                        If parser.ValidateIdentifier(fragments.ScriptTokenStream(i).Text) = False And _
    '                            fragments.ScriptTokenStream(i).TokenType.ToString() <> "SingleLineComment" And _
    '                            fragments.ScriptTokenStream(i).TokenType.ToString() <> "Variable" And _
    '                            fragments.ScriptTokenStream(i).TokenType.ToString() <> "WhiteSpace" Then
    '                            rtb.Select(fragments.ScriptTokenStream(i).Offset, fragments.ScriptTokenStream(i).Text.Length)
    '                            rtb.SelectionColor = Color.Blue
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        End If
    '    Next
    '    'MsgBox(strTokens)
    'End Sub
End Module
