'------------------------------------------------
'Name: Module ConnectionString.vb
'Function: Construct the connection strings.
'Copyright Baines 2012. All rights reserved.
'Notes:
'Modifications:
'------------------------------------------------
Imports System.Configuration
Public Class ConnectionString
    Public Shared ConnectionString As String = ""
    Public Shared Quality As Boolean = False

    'This is for SQL Authentication
    'Sub New(FileName As String, DefaultcipherText As String, _
    Public Shared Sub Init(FileName As String, DefaultcipherText As String, _
    ByVal UserId As String, _Quality As Boolean, DataSourceDevelopment As String, _
    DataSourceDevelopmentTest As String, _
    CatalogDevelopment As String, _
    DataSourceTest As String, _
    CatalogTest As String, _
    DataSourceLive As String, _
    CatalogLive As String)

        Dim DataSource As String = ""
        Dim Catalog As String = ""
        Dim strDevelopment As String = DataSourceDevelopment

        Quality = _Quality
        If My.Computer.Name.Length < DataSourceDevelopment.Length Then
            strDevelopment = DataSourceDevelopment.Substring(0, My.Computer.Name.Length)
        End If

        If My.Computer.Name = strDevelopment Then
            If Quality = True Then
                DataSource = DataSourceDevelopmentTest
            Else
                DataSource = DataSourceDevelopment
            End If
            Catalog = CatalogDevelopment

        Else
            If Quality = True Then
                DataSource = DataSourceTest
                Catalog = CatalogTest
            Else
                DataSource = DataSourceLive
                Catalog = CatalogLive
            End If
        End If

        Dim deCode As New DeEnCode(FileName, DefaultcipherText) '"cipherText.txt", "KW6xdELlM57NgAAMR3psE5sh6/RkdJ1o")
        Dim Password As String = deCode.DecryptData()
        ConnectionString = "Data Source=" & DataSource & ";Initial Catalog=" & Catalog & ";User Id=" & UserId & ";Password=" & Password & ";"
    End Sub

    'This is for Windows Authentication
    Public Shared Sub Init(_Quality As Boolean, DataSourceDevelopment As String, _
        DataSourceDevelopmentTest As String, _
        CatalogDevelopment As String, _
        DataSourceTest As String, _
        CatalogTest As String, _
        DataSourceLive As String, _
        CatalogLive As String)

        Dim DataSource As String
        Dim Catalog As String
        Quality = _Quality
        Dim strDevelopment As String = DataSourceDevelopment
        If My.Computer.Name.Length < DataSourceDevelopment.Length Then
            strDevelopment = DataSourceDevelopment.Substring(0, My.Computer.Name.Length)
        End If

        If My.Computer.Name = strDevelopment Then
            If Quality = True Then
                DataSource = DataSourceDevelopmentTest
            Else
                DataSource = DataSourceDevelopment
            End If
            Catalog = CatalogDevelopment
        Else
            If Quality = True Then
                DataSource = DataSourceTest
                Catalog = CatalogTest
            Else
                DataSource = DataSourceLive
                Catalog = CatalogLive
            End If
        End If
        ConnectionString = "Data Source=" & DataSource & ";Initial Catalog=" & Catalog & ";Integrated Security=True"
    End Sub

End Class

