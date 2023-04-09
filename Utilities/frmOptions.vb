'------------------------------------------------
'Name: Module genForm_b_app_parameter.vb.
'Function: The Options form.
'Copyright Robin Baines 2010. All rights reserved.
'Created 4/8/2010 12:00:00 AM.
'Notes: 
'Modifications: Some editing because generated for utilities.
'Allow user to delete and add set to false. Renamed frmOptions instead of frm_b_app_parameter.
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class frmOptions
    Inherits Genericform
    Friend WithEvents TheDataSet As TheDataSet
    'Protected WithEvents taParent As TheDataSetTableAdapters.b_app_parameterTableAdapter
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
    Friend WithEvents ColorDialog1 As System.Windows.Forms.ColorDialog
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Private components As System.ComponentModel.IContainer

    Private Sub InitializeComponent()
        Me.TheDataSet = New TheDataSet
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog
        Me.ListBox1 = New System.Windows.Forms.ListBox
        CType(Me.bsParent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TheDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'bsParent
        '
        Me.bsParent.DataMember = "b_app_parameter"
        Me.bsParent.DataSource = Me.TheDataSet
        '
        'TheDataSet
        '
        Me.TheDataSet.DataSetName = "TheDataSet"
        Me.TheDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(718, 329)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(120, 95)
        Me.ListBox1.TabIndex = 3
        '
        'frmOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.ClientSize = New System.Drawing.Size(1093, 438)
        Me.Controls.Add(Me.ListBox1)
        Me.Name = "frmOptions"
        Me.Controls.SetChildIndex(Me.ListBox1, 0)
        CType(Me.bsParent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TheDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#Region "New"
    'This for the designer?!
    Public Sub New()
        MyBase.New()
        InitializeComponent()
        Me.taParent = New TheDataSetTableAdapters.b_app_parameterTableAdapter
    End Sub
    Public Sub New(ByVal tsb As ToolStripItem _
    , ByVal strSecurityName As String, ByVal _MainDefs As MainDefinitions, _
    ByVal Fields As Dictionary(Of String, String), ByVal _blnFullSize As Boolean, ByVal _bRO As Boolean)
        MyBase.New(tsb, strSecurityName, _MainDefs, Fields, _blnFullSize, _bRO)

        InitializeComponent()
        Me.taParent = New TheDataSetTableAdapters.b_app_parameterTableAdapter

        vParent = New b_app_parameter(MainDefs.strGetTableText(strSecurityName), Me.bsParent, Me.dgParent, _
        taParent, _
        Me.TheDataSet, _
        Me._components, _
        _MainDefs, blnRO, True, Me.Controls, Me)
        vParent.CreateFilterBoxes(Me.Controls)
        iInitialFormHeight = 1022 + 60
    End Sub
#End Region
#Region "Load"
    Protected Overrides Sub FillTableAdapter()
        vParent.RefreshCombos()
        vParent.StoreRowIndexWithFocus()
        Me.taParent.Fill(Me.TheDataSet.b_app_parameter)
        vParent.ResetFocusRow()

        'Set the filter if necessary.
        vParent.ColumnDoubleClick(FilterFields)
        dgParent.AllowUserToAddRows = False
        dgParent.AllowUserToDeleteRows = False
        'Me.ColorDialog1.ShowDialog()
        Dim cC As Color
        Dim i As Integer
        For i = 1 To 1500
            'System.Drawing.KnownColor.wind()
            Try
                cC = Color.FromKnownColor(i)
                'Return cC.name as integer i as string if not a recognised KnownColor; ie outside the range.
                If cC = Nothing Or cC.Name = i.ToString() Then
                    Exit For
                End If
                Me.ListBox1.Items.Add(cC.Name)

            Catch ex As Exception
                Exit For
            End Try

        Next
        ListBox1.Sorted = True


        'ResizeGrid(dgParent)
    End Sub
    Protected Overrides Sub RefreshCombos()
        vParent.RefreshCombos()
    End Sub
#End Region
End Class
