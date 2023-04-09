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
Imports System.Configuration.CommaDelimitedStringCollection
Public Class SQLBrowser
#Region "new"

    Dim TheNodes As New List(Of Node)()
    Dim theViews As New Dictionary(Of String, AView)
    Dim theTables As New Dictionary(Of String, String)
    Dim theProcs As New Dictionary(Of String, AView)
    Dim ListOfTCs As New List(Of TabControl)()

    Dim Server As Server = New Server(TestAppMain.GetDataSourceLive) '"BAINESLENOVO")
    Dim dbs As DatabaseCollection = Server.Databases
    Dim db As Database = dbs(TestAppMain.GetCatalogLive())

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        Init()
    End Sub

    Public Sub New(ByVal tsb As ToolStripItem _
              , ByVal strSecurityName As String, ByVal _MainDefs As MainDefinitions)

        MyBase.New(tsb, strSecurityName, _MainDefs)
        InitializeComponent()
        Init()
    End Sub

    Private Sub MakeTabPage(TC As TabControl)
        Dim TP = New System.Windows.Forms.TabPage()
        Dim rt As RichTextBox
        rt = New System.Windows.Forms.RichTextBox()
        rt.Location = New System.Drawing.Point(3, 8)
        rt.Name = "RichTextBox" + TabIndex.ToString
        rt.Size = New System.Drawing.Size(400, 496)
        rt.TabIndex = TabIndex
        rt.Text = ""
        rt.Font = New Font("Consolas", 10)
        rt.WordWrap = False
        rt.Dock = DockStyle.Fill
        AddHandler rt.TextChanged, AddressOf rtb_TextChanged
        TP.Controls.Add(rt)
        TP.Location = New System.Drawing.Point(4, 24)
        TP.Name = "TabPage1"
        TP.Padding = New System.Windows.Forms.Padding(3)
        TP.Size = New System.Drawing.Size(244, 218)
        TP.TabIndex = 0
        TP.Text = "TabPage1"
        TP.UseVisualStyleBackColor = True
        TC.Controls.Add(TP)
    End Sub

    Private Sub DropTabPages(TC As TabControl)

    End Sub

    Private Sub MakeTheTabControls(TabIndex As Integer, gb As GroupBox)

        Dim TC = New System.Windows.Forms.TabControl()
        gb.Controls.Add(TC)
        MakeTabPage(TC)
        ListOfTCs.Add(TC)

    End Sub

    Private Sub Init()
        'TVUses.CheckBoxes = True
        'TVUsedBy.CheckBoxes = True
        MakeTheTabControls(10, gbUses)
        MakeTheTabControls(20, gbUses)
        MakeTheTabControls(30, gbUses)

        MakeTheTabControls(10, gbUsedBy)
        MakeTheTabControls(20, gbUsedBy)
        MakeTheTabControls(30, gbUsedBy)

        Me.SwitchOffPrintDetail()
        Me.SwitchOffPrint()
        Me.SwitchOffUpdate()

    End Sub
#End Region
#Region "Load"

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Server.SetDefaultInitFields(GetType(View), {"IsSystemObject", "Text"})
        Server.SetDefaultInitFields(GetType(Table), {"IsSystemObject", "Schema"})
        Server.SetDefaultInitFields(GetType(StoredProcedure), {"IsSystemObject", "Text"})
        GetViewsProcsAndTables(db)
        ConnectedViews(db)
        Server.ConnectionContext.Disconnect()
        ObjectUses(TVUses)
        ObjectUsedBy(TVUsedBy)
    End Sub
#End Region
#Region "Uses"
    Private Sub AddTreeUsesNodes(nodes As TreeNodeCollection, nValue As String) 'As Boolean
        For Each n As Node In TheNodes
            If n.Key = nValue Then
                Dim iNewIndex = nodes.Add(New TreeNode(n.Value))
                nodes(iNewIndex).Name = n.Value
                If n.IsInLine Then
                    nodes(iNewIndex).BackColor = Color.Aqua
                End If
                AddTreeUsesNodes(nodes(iNewIndex).Nodes, n.Value)
            End If
        Next
    End Sub

    Sub ObjectUses(Tv As TreeView)

        Tv.BeginUpdate()

        ' Clear the TreeView each time the method is called.
        Tv.Nodes.Clear()
        Dim iTablesIndex = Tv.Nodes.Add(New TreeNode("Tables"))
        Tv.Nodes(iTablesIndex).Name = "Tables"
        Dim iViewsIndex = Tv.Nodes.Add(New TreeNode("Views"))
        Tv.Nodes(iViewsIndex).Name = "Views"
        Dim iProcsIndex = Tv.Nodes.Add(New TreeNode("Procs"))
        Tv.Nodes(iProcsIndex).Name = "Procs"

        For Each t As KeyValuePair(Of String, String) In theTables
            Dim iIndex = Tv.Nodes(iTablesIndex).Nodes.Add(New TreeNode(t.Key))
            Tv.Nodes(iTablesIndex).Nodes(iIndex).Name = t.Key
        Next
        For Each t As KeyValuePair(Of String, AView) In theViews
            Dim iIndex = Tv.Nodes(iViewsIndex).Nodes.Add(New TreeNode(t.Key))
            Tv.Nodes(iViewsIndex).Nodes(iIndex).Name = t.Key
            For Each n As Node In TheNodes
                If n.Key = t.Key Then
                    If n.Key = "v_aid_result_last_session" Then
                        Debug.Print("v_aid_result_last_session")
                    End If
                    'add the nodes to the treeview.
                    Dim iIndex2 = Tv.Nodes(iViewsIndex).Nodes(iIndex).Nodes.Add(New TreeNode(n.Value))
                    Tv.Nodes(iViewsIndex).Nodes(iIndex).Nodes(iIndex2).Name = n.Value
                    If n.IsInLine Then
                        Tv.Nodes(iViewsIndex).Nodes(iIndex).Nodes(iIndex2).BackColor = Color.Aqua
                    End If

                    'if the view does not use other views or tables give it another colour.
                    If n.Value = "" Then
                        Tv.Nodes(iViewsIndex).Nodes(iIndex).Nodes(iIndex2).BackColor = Color.DarkGray
                    End If
                    AddTreeUsesNodes(Tv.Nodes(iViewsIndex).Nodes(iIndex).Nodes(iIndex2).Nodes, n.Value)
                End If
            Next
        Next
        For Each t As KeyValuePair(Of String, AView) In theProcs
            Dim iIndex = Tv.Nodes(iProcsIndex).Nodes.Add(New TreeNode(t.Key))
            Tv.Nodes(iProcsIndex).Nodes(iIndex).Name = t.Key
            For Each n As Node In TheNodes
                If n.Key = t.Key Then
                    'add the nodes to the treeview.
                    Dim iIndex2 = Tv.Nodes(iProcsIndex).Nodes(iIndex).Nodes.Add(New TreeNode(n.Value))
                    Tv.Nodes(iProcsIndex).Nodes(iIndex).Nodes(iIndex2).Name = n.Value
                    If n.IsInLine Then
                        Tv.Nodes(iProcsIndex).Nodes(iIndex).Nodes(iIndex2).BackColor = Color.Aqua
                    End If

                    'if the view does not use other views or tables give it another colour.
                    If n.Value = "" Then
                        Tv.Nodes(iProcsIndex).Nodes(iIndex).Nodes(iIndex2).BackColor = Color.DarkGray
                    End If
                    AddTreeUsesNodes(Tv.Nodes(iProcsIndex).Nodes(iIndex).Nodes(iIndex2).Nodes, n.Value)
                End If
            Next
        Next

        ' Tv.Sort()
        Tv.EndUpdate()
    End Sub
#End Region

#Region "UsedBy"
    Private Sub AddTreeUsedByNodes(nodes As TreeNodeCollection, nValue As String) 'As Boolean
        For Each n As Node In TheNodes
            If n.Value = nValue Then
                Dim iNewIndex = nodes.Add(New TreeNode(n.Key))
                nodes(iNewIndex).Name = n.Key
                'If n.IsInLine Then
                '    nodes(iNewIndex).BackColor = Color.Aqua
                'End If
                AddTreeUsedByNodes(nodes(iNewIndex).Nodes, n.Key)
            End If
        Next
    End Sub

    Sub ObjectUsedBy(Tv As TreeView)

        Tv.BeginUpdate()

        ' Clear the TreeView each time the method is called.
        Tv.Nodes.Clear()
        Dim iTablesIndex = Tv.Nodes.Add(New TreeNode("Tables"))
        Tv.Nodes(iTablesIndex).Name = "Tables"
        Dim iViewsIndex = Tv.Nodes.Add(New TreeNode("Views"))
        Tv.Nodes(iViewsIndex).Name = "Views"
        'Dim iProcsIndex = Tv.Nodes.Add(New TreeNode("Procs"))
        'Tv.Nodes(iProcsIndex).Name = "Procs"

        For Each t As KeyValuePair(Of String, String) In theTables
            Dim iIndex = Tv.Nodes(iTablesIndex).Nodes.Add(New TreeNode(t.Key))
            Tv.Nodes(iTablesIndex).Nodes(iIndex).Name = t.Key
        Next

        For Each t As KeyValuePair(Of String, AView) In theViews
            Dim iIndex = Tv.Nodes(iViewsIndex).Nodes.Add(New TreeNode(t.Key))
            Tv.Nodes(iViewsIndex).Nodes(iIndex).Name = t.Key
            For Each n As Node In TheNodes
                If n.Value = t.Key Then
                    'add the nodes to the treeview.
                    Dim iIndex2 = Tv.Nodes(iViewsIndex).Nodes(iIndex).Nodes.Add(New TreeNode(n.Key))
                    Tv.Nodes(iViewsIndex).Nodes(iIndex).Nodes(iIndex2).Name = n.Key
                    'If n.IsInLine Then
                    '    Tv.Nodes(iViewsIndex).Nodes(iIndex).Nodes(iIndex2).BackColor = Color.Aqua
                    'End If

                    'if the view does not use other views or tables give it another colour.
                    'If n.Value = "" Then
                    '    Tv.Nodes(iViewsIndex).Nodes(iIndex).Nodes(iIndex2).BackColor = Color.DarkGray
                    'End If
                    AddTreeUsedByNodes(Tv.Nodes(iViewsIndex).Nodes(iIndex).Nodes(iIndex2).Nodes, n.Key)
                End If
            Next
        Next

        'For Each n As Node In TheNodes
        '    If n.Key = "v_activity" Then
        '        Debug.Print("v_activity")
        '    End If

        '    'For Each t As KeyValuePair(Of String, String) In theTables
        '    '    Dim iIndex = Tv.Nodes(iTablesIndex).Nodes.Add(New TreeNode(t.Key))
        '    '    Tv.Nodes(iTablesIndex).Nodes(iIndex).Name = t.Key
        '    'Next

        '    'add the nodes to the treeview.
        '    If Tv.Nodes(iViewsIndex).Nodes.ContainsKey(n.Key) = False Then
        '        Dim iIndex = Tv.Nodes(iViewsIndex).Nodes.Add(New TreeNode(n.Key))
        '        Tv.Nodes(iViewsIndex).Nodes(iIndex).Name = n.Key
        '        For Each n2 As Node In TheNodes
        '            If n2.Value = n.Key Then
        '                Dim iIndex2 = Tv.Nodes(iViewsIndex).Nodes(iIndex).Nodes.Add(New TreeNode(n2.Key))
        '                Tv.Nodes(iViewsIndex).Nodes(iIndex).Nodes(iIndex2).Name = n2.Key
        '                AddTreeUsedByNodes(Tv.Nodes(iViewsIndex).Nodes(iIndex).Nodes(iIndex2).Nodes, n2.Key)
        '            End If
        '        Next
        '    End If
        'Next
        'Tv.Sort()
        Tv.EndUpdate()
    End Sub
#End Region

#Region "AddTheNodes"

    Function RemoveComments(str As String) As String
        Dim charsToTrim() As Char = {" ", vbLf, vbTab}
        Dim fields() As String = str.Split(New Char() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
        Dim strRet As String = ""
        ' Dim iCount As Integer = 0
        For Each strT As String In fields
            strT = strT.Trim(charsToTrim)
            Dim iComment As Integer = strT.IndexOf("--")
            If iComment > -1 Then
                strT = strT.Substring(0, iComment)
            End If
            strRet += strT
            strRet += vbCrLf
        Next
        Return strRet
    End Function

    Private Sub AddTheNodes(file As System.IO.StreamWriter, Output_Name As String, strTable As String, blnInline As Boolean)
        Dim fields() As String = strTable.Split(New Char() {" ", ",", "\r", "\n", vbCrLf}, StringSplitOptions.RemoveEmptyEntries)

        For Each strT As String In fields
            strT = strT.Replace(vbCr, "").Replace(vbLf, "").Replace("[", "").Replace("]", "")

            'If there is no schema add the default schema.
            strT = strT.ToUpper()
            If strT.Contains(".") = False Then
                strT = "DBO." + strT
            End If

            'add views.
            If theViews.ContainsKey(strT) Then

                'Just a list of 'a' uses 'b'
                Dim n As New Node(Output_Name, strT, False, False, True, blnInline)
                If Not TheNodes.Contains(n) Then
                    If blnInline Then
                        file.WriteLine(Output_Name + " uses view inline " + strT)
                    Else
                        file.WriteLine(Output_Name + " uses view " + strT)
                    End If
                    TheNodes.Add(n)
                End If
            Else

                'add tables
                If theTables.ContainsKey(strT) Then
                    Dim n As New Node(Output_Name, strT, False, True, False, blnInline)
                    If Not TheNodes.Contains(n) Then
                        If blnInline Then
                            file.WriteLine(Output_Name + " uses table inline " + strT)
                        Else
                            file.WriteLine(Output_Name + " uses table " + strT)
                        End If
                        TheNodes.Add(n)
                    End If
                End If
            End If
        Next

    End Sub

    Private Function GetTableScript(db As Database, strT As String) As String
        Dim strScript As String = ""
        For Each t As Table In db.Tables
            If t.IsSystemObject = False Then
                If strT = (t.Schema + "." + t.Name.Replace(" ", "_").Replace("$", "_")).ToUpper() Then
                    Dim scriptOptions = New ScriptingOptions()
                    scriptOptions.ScriptDrops = False
                    scriptOptions.IncludeIfNotExists = False
                    scriptOptions.AnsiPadding = False
                    scriptOptions.DriForeignKeys = True
                    scriptOptions.DriIndexes = True

                    ' scriptOptions.
                    'get the table defininition (takes a long time!)
                    Dim tableScripts As Specialized.StringCollection = t.Script(scriptOptions)

                    For Each str As String In tableScripts
                        If Not str.Contains("SET ANSI_NULLS") Then
                            If Not str.Contains("QUOTED_IDENTIFIER") Then
                                strScript += str
                            End If
                        End If
                    Next
                    Return strScript
                End If
            End If
        Next
        Return strScript
    End Function

    ''' <summary>
    ''' Get the views and tables.
    ''' </summary>
    ''' <param name="db"></param>
    Sub GetViewsProcsAndTables(db As Database)
        For Each t As View In db.Views
            If t.IsSystemObject = False Then
                Dim Output_Name As String = t.Schema + "." + t.Name.Replace(" ", "_").Replace("$", "_")
                Output_Name = Output_Name.ToUpper()
                If Not theViews.ContainsKey(Output_Name.ToUpper()) Then
                    theViews.Add(Output_Name.ToUpper(), New AView(t.TextHeader, t.TextBody))
                End If
            End If
        Next
        For Each t As Table In db.Tables
            If t.IsSystemObject = False Then
                Dim Output_Name As String = t.Schema + "." + t.Name.Replace(" ", "_").Replace("$", "_")
                If Not theTables.ContainsKey(Output_Name.ToUpper()) Then
                    theTables.Add(Output_Name.ToUpper(), "")
                End If
            End If
        Next
        For Each t As StoredProcedure In db.StoredProcedures
            If t.IsSystemObject = False Then
                Dim Output_Name As String = t.Schema + "." + t.Name.Replace(" ", "_").Replace("$", "_")
                If Not theProcs.ContainsKey(Output_Name.ToUpper()) Then
                    theProcs.Add(Output_Name.ToUpper(), New AView(t.TextHeader, t.TextBody))
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Store which views and tables a view uses and write this list of view uses view/table to a file.
    ''' </summary>
    ''' <param name="db"></param>
    Sub ConnectedViews(db As Database)
        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter("c:\temp\ViewUses.txt", False)
        For Each t As KeyValuePair(Of String, AView) In theViews
            Dim Output_Name As String = t.Key
            If Output_Name = "DBO.v_lot_serial_result".ToUpper() Then
                Debug.Print(Output_Name)
            End If

            'Split the SQL into the fields section and the FROM/JOIN section.
            'The splitting needs to be improved because of inline (SELECT .. FROM) in the FROM/JOIN part of the VIEW.
            Dim strSQL As String = t.Value.TextBody.ToUpper()
            strSQL = RemoveComments(strSQL)
            Dim iFrom As Integer = strSQL.LastIndexOf("FROM")
            If iFrom > -1 Then

                'add the nodes from the FROM/JOIN section of the view.
                AddTheNodes(file, Output_Name, strSQL.Substring(iFrom), False)

                ''add the nodes from inline use of a Tables and Views.
                Dim strTe As String = strSQL.Substring(0, iFrom)
                Dim iFrom2 = strTe.IndexOf("FROM")
                If iFrom2 > -1 Then
                    AddTheNodes(file, Output_Name, strTe, True)
                End If
            End If

            'Some views do not use other views or tables. Add them to TheNodes.
            Dim blnIsNode As Boolean = False
            For Each n As Node In TheNodes
                If n.Key = t.Key Then
                    blnIsNode = True
                    Exit For
                End If
            Next n
            If blnIsNode = False Then
                Dim n As New Node(Output_Name, "", False, False, True, False)
                If Not TheNodes.Contains(n) Then
                    file.WriteLine(Output_Name + " does not use application views or tables.")
                    TheNodes.Add(n)
                End If
            End If
        Next t

        For Each t As KeyValuePair(Of String, AView) In theProcs
            Dim strSQL As String = t.Value.TextBody.ToUpper()
            strSQL = RemoveComments(strSQL)
            'add the nodes from the FROM/JOIN section of the view.
            AddTheNodes(file, t.Key, strSQL, False)
        Next

        file.Close()
    End Sub

#End Region


#Region "AfterSelect"

    Private Function GetLevel(n As TreeNode) As Integer
        Dim iLevel = n.Level - 1
        If iLevel > 2 Then
            iLevel = 2
        End If
        Return iLevel
    End Function

    Private Sub ShowAllAtLevel(e As TreeViewEventArgs, iOffset As Integer)


        If e.Node.Level > 0 Then

            'show the sources of the selected node.
            ShowALevel(e.Node, 0, iOffset)

            'then load the sources of children into the TabPages of the next TabControl
            Dim iTabPage As Integer = 0
            Dim TC As TabControl = Nothing
            For Each n As TreeNode In e.Node.Nodes
                TC = ListOfTCs(iOffset + GetLevel(n))
                If TC.TabPages.Count < (iTabPage + 1) Then
                    MakeTabPage(ListOfTCs(iOffset + GetLevel(n)))
                End If
                ShowALevel(n, iTabPage, iOffset)
                iTabPage += 1
            Next
            If Not TC Is Nothing Then
                Dim iTpages = TC.TabPages.Count
                While iTpages > iTabPage
                    TC.TabPages(iTpages - 1).Dispose()
                    iTpages -= 1
                End While
            End If
        End If
    End Sub

    Private Sub ShowALevel(n As TreeNode, iTabPage As Integer, iOffset As Integer)
        Dim strText As String = ""
        If theViews.ContainsKey(n.Name) Then
            Dim AView As AView = theViews.GetValueOrDefault(n.Name)
            If Not AView Is Nothing Then
                strText = AView.TextHeader + AView.TextBody
            End If
        Else
            If theTables.ContainsKey(n.Name) Then
                Dim ATable As String = theTables.GetValueOrDefault(n.Name)
                If ATable = "" Then
                    ATable = GetTableScript(db, n.Name)
                    theTables(n.Name) = ATable
                End If
                strText = ATable
            Else
                If theProcs.ContainsKey(n.Name) Then
                    Dim AProc As AView = theProcs.GetValueOrDefault(n.Name)
                    If Not AProc Is Nothing Then
                        strText = AProc.TextHeader + AProc.TextBody
                    End If
                End If
            End If
        End If
        ListOfTCs(iOffset + GetLevel(n)).TabPages(iTabPage).Controls(0).Text = strText
        ListOfTCs(iOffset + GetLevel(n)).TabPages(iTabPage).Text = n.Name.Replace("DBO.", "")
    End Sub

    Private Sub TVUses_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TVUses.AfterSelect
        ShowAllAtLevel(e, 0)
        ' ShowALevel(e.Node, 0)
    End Sub

    Private Sub TVUsedBy_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TVUsedBy.AfterSelect
        ShowAllAtLevel(e, 3)
        'ShowALevel(e.Node, 3)
    End Sub

#End Region
#Region "Scroll"
    Private Sub ReSizeC()
        If Not gbUses Is Nothing And Not gbUsedBy Is Nothing Then
            gbUses.Size = New Point(Me.ClientRectangle.Width - gbUses.Location.X - 20, (Me.ClientRectangle.Height - gbUses.Location.Y - 10) / 2)
            TVUses.Size = New Point(TVUses.Width, gbUses.Size.Height - 30)

            gbUsedBy.Location = New Point(gbUses.Location.X, gbUses.Height + gbUses.Location.Y + 5)
            gbUsedBy.Size = New Point(Me.ClientRectangle.Width - gbUses.Location.X - 20, gbUses.Size.Height) '(Me.ClientRectangle.Height - gbUsedBy.Location.Y - 10))
            TVUsedBy.Size = New Point(TVUsedBy.Width, gbUsedBy.Size.Height - 30)

            Dim iWidth = (gbUses.Width - gbUses.Location.X - TVUses.Width - 8) / 3
            Dim iX = TVUses.Location.X + TVUses.Width + 5
            Dim iY = TVUses.Location.Y
            Dim iCount As Integer = 0
            ' For Each rt As RichTextBox In ListOfRTFs
            For Each rt As TabControl In ListOfTCs
                If (iCount = 3) Then
                    iCount = 0
                End If
                If Not rt Is Nothing Then
                    rt.Location = New Point(iX + iWidth * iCount, iY)
                    rt.Size = New Point(iWidth, gbUses.Size.Height - 30)
                    iY = rt.Location.Y
                    iCount += 1
                End If
            Next
        End If

    End Sub

    Protected Overrides Sub frm_Layout(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles MyBase.Layout
        MyBase.frm_Layout(sender, e)
        ReSizeC()
    End Sub
#End Region

    Public SQLParser As New clParseSQL()

#Region "parse_sql"
    Private Sub rtb_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        Try
            SQLParser.ParseRTB(sender)
        Catch ex As Exception

        End Try
    End Sub
#End Region
#Region "Search"

    Private Sub AfterExpand(e As TreeViewEventArgs)
        If Not e.Node.Parent Is Nothing Then
            If e.Node.Parent.Name = "Views" Or e.Node.Parent.Name = "Procs" Or e.Node.Parent.Name = "Tables" Then
                e.Node.ExpandAll()
            End If
        End If
    End Sub

    Private Sub TVUses_AfterExpand(sender As Object, e As TreeViewEventArgs) Handles TVUses.AfterExpand
        AfterExpand(e)
    End Sub
    Private Sub TVUsedBy_AfterExpand(sender As Object, e As TreeViewEventArgs) Handles TVUsedBy.AfterExpand
        AfterExpand(e)
    End Sub

    'Private Sub TVUses_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TVUses.AfterSelect
    '    If e.Node.Parent Is Nothing Then
    '        e.Node.ExpandAll()
    '    End If
    'End Sub
    'Private Sub TVUsedBy_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TVUsedBy.AfterSelect
    '    If e.Node.Parent Is Nothing Then
    '        e.Node.ExpandAll()
    '    End If
    'End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim iUsesIndex As Integer = -1
        TVUses.CollapseAll()
        TVUsedBy.CollapseAll()

        'create an upper case version of the search text preceded by 'DBO.' if not provided by the user.
        Dim strSearchText As String = tbSchema.Text.ToUpper()
        If strSearchText.Length > 0 And Not strSearchText.Contains(".") Then
            strSearchText = strSearchText + "."
        End If
        If strSearchText.Length = 0 Then
            strSearchText = "DBO."
        End If
        strSearchText += tbSearch.Text.ToUpper()

        Dim nFound As TreeNode = Nothing

        For Each n As TreeNode In TVUses.Nodes
            For Each n2 As TreeNode In n.Nodes
                If n2.Name.Contains(strSearchText) Then
                    iUsesIndex = n2.Index
                    nFound = n
                    Exit For
                End If
            Next
        Next

        If iUsesIndex > -1 And Not nFound Is Nothing Then
            Dim iUsedByIndex As Integer = -1
            For Each n As TreeNode In TVUsedBy.Nodes
                If n.Name = nFound.Nodes(iUsesIndex).Name Then
                    iUsedByIndex = n.Index
                    Exit For
                End If
            Next

            If iUsedByIndex > -1 Then
                TVUsedBy.Focus()
                TVUsedBy.Nodes(iUsedByIndex).Expand() 'Fires AfterExpand Which does the .ExpandAll()
                TVUsedBy.SelectedNode = TVUsedBy.Nodes(iUsedByIndex)
            End If

            TVUses.Focus()
            nFound.Nodes(iUsesIndex).Expand() 'Fires AfterExpand Which does the .ExpandAll()

            'Select after the Expand to ensure the root node is shown.
            TVUses.SelectedNode = nFound.Nodes(iUsesIndex)
        End If
    End Sub
#End Region
End Class
Public Class AView
    Public Sub New(_TextHeader As String, _TextBody As String)
        TextHeader = _TextHeader
        TextBody = _TextBody
    End Sub
    Public TextHeader As String
    Public TextBody As String
End Class

Public Class Node
    Implements IEquatable(Of Node)
    Public Sub New(_Key As String, _Value As String, _IsDataSet As Boolean, _IsTable As Boolean, _IsView As Boolean, _IsInLine As Boolean)
        Key = _Key
        Value = _Value
        IsDataSet = _IsDataSet
        IsTable = _IsTable
        IsView = _IsView
        IsInLine = _IsInLine
    End Sub

    Public Function ContainsValue(kvp As Node, org_kvp As Node) As Boolean
        If (Key = kvp.Value And Key <> org_kvp.Key) Then Return True
        Return False
    End Function


    ''' <summary>
    ''' 'Used by Contains(n Node)
    ''' </summary>
    ''' <param name="other"></param>
    ''' <returns></returns>
    Public Overloads Function Equals(ByVal other As Node) _
            As Boolean Implements IEquatable(Of Node).Equals
        If Me.Key = other.Key And Me.Value = other.Value Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Key As String
    Public Value As String
    Public IsDataSet As Boolean
    Public IsTable As Boolean
    Public IsView As Boolean
    Public IsInLine As Boolean

End Class
