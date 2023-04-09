'------------------------------------------------
'Name: Module 
'Function: 
'Copyright Baines 2013. All rights reserved.
'Notes: 
'Modifications: 
'PROBLEM WITH .NET Core - Using system.diagnostics in App.config. Remove the system.diagnostics section in CoreTestApp.dll.config.
'------------------------------------------------
Imports Utilities
Imports Microsoft.SqlServer.Management.Common
Imports Microsoft.SqlServer.Management.Smo
Imports System
Imports System.Configuration
Imports System.Data.SqlClient


Public Class HTMLView
#Region "new"
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        webBrowser1.AllowWebBrowserDrop = False
        webBrowser1.DocumentText = "<HTML> AView.TextHeader + AView.TextBody <\HTML> "
        'webBrowser1.GoHome()
    End Sub
#End Region
    Public Sub ShowDoc(str As String)
        webBrowser1.DocumentText = str

    End Sub

    'Private Sub webBrowser1_Navigating(
    '    ByVal sender As Object, ByVal e As WebBrowserNavigatingEventArgs) _
    '    Handles webBrowser1.Navigating

    '    Dim document As System.Windows.Forms.HtmlDocument =
    '        webBrowser1.Document
    '    If document IsNot Nothing And
    '        document.All("userName") IsNot Nothing And
    '        String.IsNullOrEmpty(
    '        document.All("userName").GetAttribute("value")) Then

    '        e.Cancel = True
    '        MsgBox("You must enter your name before you can navigate to " &
    '            e.Url.ToString())
    '    End If

    'End Sub
End Class