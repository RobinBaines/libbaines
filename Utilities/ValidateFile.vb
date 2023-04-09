'------------------------------------------------------------------------
'Name: Validatefile.vb.
'Function: Validate an XML file against an xsd.
'Copyright Robin Baines 2005. All rights reserved.
'By: Robin Baines. 
'Creation Date: 29/06/2005
'Modifications: 20090615 Added overload for CheckXMLFile.
'------------------------------------------------------------------------
Imports System.Xml         'for XmlTextReader and XmlValidatingReader.
Imports System.Xml.Schema  'for XmlSchemaCollection. 
Imports System.Xml.XPath
Imports System.Xml.Xsl

Public Class ValidateFile
    Private strErrorText As String
    Public ReadOnly Property ErrorText() As String
        Get
            Return strErrorText
        End Get
    End Property

    Private blnIsValidXMLFile As Boolean
    Public Property IsValidXMLFile() As Boolean
        Get
            Return blnIsValidXMLFile
        End Get
        Set(ByVal value As Boolean)
            blnIsValidXMLFile = value
        End Set
    End Property

    Public Sub New()
        strErrorText = ""
        blnIsValidXMLFile = False
    End Sub

    Protected Overridable Sub ValidationEventHandle(ByVal sender As Object, ByVal args As ValidationEventArgs)
        blnIsValidXMLFile = False
        Console.WriteLine(ControlChars.CrLf & ControlChars.Tab & "Validation error: " & args.Message)
        strErrorText = strErrorText & "Validation error: " & args.Message
        'Do not add the error text as an XML comment as the file may not be xml.

    End Sub

    '20090615 Added this overload for CheckXMLFile.
    Public Function CheckXMLFile(ByVal FileName As String, ByVal strSchema As String, ByVal strNameSpace As String) As Boolean
        Return CheckXMLFile(FileName, strSchema, strNameSpace, ValidationType.Schema)
    End Function

    Public Function CheckXMLFile(ByVal FileName As String, ByVal strSchema As String, ByVal strNameSpace As String, ByVal ValidationType As ValidationType) As Boolean

        'Check the XML file against the schema using the namespace. 
        'ValidationEventHandle is called if there is an error.
        Dim settings As New XmlReaderSettings()

        'Set up for inline schema checking.
        settings.ValidationType = ValidationType.Schema
        settings.ValidationFlags = settings.ValidationFlags Or XmlSchemaValidationFlags.ProcessInlineSchema
        settings.Schemas.Add(strNameSpace, strSchema)
        AddHandler settings.ValidationEventHandler, AddressOf ValidationEventHandle
        Dim reader As XmlReader = XmlReader.Create(FileName, settings)

        If reader.EOF Then
            blnIsValidXMLFile = False
        Else
            blnIsValidXMLFile = True
        End If

        'Parse the file.
        Try
            reader.MoveToContent()
            While reader.Read()
                'Select Case reader.NodeType
                '    Case XmlNodeType.Element
                '        Console.Write("<{0}>", reader.Name)
                '    Case XmlNodeType.Text
                '        Console.Write(reader.Value)
                '    Case XmlNodeType.CDATA
                '        Console.Write("<![CDATA[{0}]]>", reader.Value)
                '    Case XmlNodeType.ProcessingInstruction
                '        Console.Write("<?{0} {1}?>", reader.Name, reader.Value)
                '    Case XmlNodeType.Comment
                '        Console.Write("<!--{0}-->", reader.Value)
                '    Case XmlNodeType.XmlDeclaration
                '        Console.Write("<?xml version='1.0'?>")
                '    Case XmlNodeType.Document
                '    Case XmlNodeType.DocumentType
                '        Console.Write("<!DOCTYPE {0} [{1}]", reader.Name, reader.Value)
                '    Case XmlNodeType.EntityReference
                '        Console.Write(reader.Name)
                '    Case XmlNodeType.EndElement
                '        Console.Write("</{0}>", reader.Name)
                'End Select
            End While
        Catch ex As Exception
            blnIsValidXMLFile = False
            Console.WriteLine(ControlChars.CrLf & ControlChars.Tab & "Not an XML file: " & ex.ToString)
            strErrorText = strErrorText & "Not an XML file: " & ex.ToString & "."
        Finally
            If Not (reader Is Nothing) Then
                reader.Close()
            End If
        End Try

        'The validation check will only be possible if Siemens adhere to xsd file.
        'CHECK. Sept 2005. Xsd file is setup but ignore valid for now.
        Return blnIsValidXMLFile
    End Function
End Class
