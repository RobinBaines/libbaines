'------------------------------------------------
'Name: Module frmAppParameters.vb
'Function: Application paramters form.
'Copyright Robin Baines 2010. All rights reserved.
'Notes: 
'Modifications: 
'------------------------------------------------
Imports System.Windows.Forms
Imports System.Drawing

Public Class frmAppParameters

    Dim vParent1 As b_app_parameter
    Dim vParent2 As b_app_color

#Region "New"
    Public Sub New()

        MyBase.New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

    End Sub
    Public Sub New(ByVal tsb As ToolStripItem _
               , ByVal strSecurityName As String, ByVal _MainDefs As MainDefinitions)

        MyBase.New(tsb, strSecurityName, _MainDefs)
        InitializeComponent()
        vParent1 = New b_app_parameter(strSecurityName, B_app_parameterBindingSource, dgParent1, B_app_parameterTableAdapter, _
            Me.TheDataSet, _
            Me.components, _
            MainDefs, blnRO, True, Controls, Me)

        vParent2 = New b_app_color(strSecurityName, B_app_colorBindingSource, dgParent2, B_app_colorTableAdapter, _
           Me.TheDataSet, _
           Me.components, _
           MainDefs, blnRO, True, Controls, Me)

        SetBindingNavigatorSource(B_app_parameterBindingSource)
        vParent1.CreateFilterBoxes(Me.Controls)
        vParent2.CreateFilterBoxes(Me.Controls)

        'Me.SwitchOffPrintDetail()
        SetPrintDetail(statics.get_txt_header("Print Colors"))
        SwitchOffNavigator()
        'Me.SwitchOffRefresh()
        'Me.SwitchOffUpdate()
        dgParent2.SelectionMode = DataGridViewSelectionMode.CellSelect  '.RowHeaderSelect
        'dgChild1.DefaultCellStyle.SelectionBackColor = Color.Transparent

    End Sub

#End Region
#Region "Load"
    Protected Overrides Sub frmLoad(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.frmLoad(sender, e)
        vParent1.Adjustcolumns(True)
        vParent1.AdjustFilterBoxes()
        vParent2.Adjustcolumns(True)
        vParent2.AdjustFilterBoxes()
        ' vParent2.AdjustPosition(vParent1)
        dgParent1.AllowUserToAddRows = False
        dgParent1.AllowUserToDeleteRows = False
        dgParent2.AllowUserToAddRows = False
        dgParent2.AllowUserToDeleteRows = False

        '20101116 Turn off the cursor select for the color field so that the colour is not changed by selection.
        vParent2.dgColor.DefaultCellStyle.SelectionBackColor = Color.Transparent
        'vParent1.dgResize(Me.Size.Width, Me.GetScrollState(ScrollStateVScrollVisible))
        'vParent2.dgResize(Me.Size.Width, Me.GetScrollState(ScrollStateVScrollVisible))

        blnAllowUpdate = True
        FillTableAdapter()

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
                If cC <> Color.Transparent Then
                    'The text box does not support some colours like Transparent so filter these out with a test.
                    Try
                        TextBox1.BackColor = cC
                        Me.ListBox1.Items.Add(cC.Name)

                    Catch ex As Exception
                        'MsgBox(cC.ToString())
                    End Try
                End If

            Catch ex As Exception
                Exit For
            End Try

        Next
        ListBox1.Sorted = True
        ListBox1.Height = dgParent2.Height
        Dim p As Point = ListBox1.Location
        p.X = dgParent2.Location.X + dgParent2.Width + 10
        ListBox1.Location = p
        p = TextBox1.Location
        p.X = ListBox1.Location.X + ListBox1.Width + 10
        Me.TextBox1.Location = p

    End Sub
    Protected Overrides Sub FillTableAdapter()
        MyBase.FillTableAdapter()
        If blnAllowUpdate = True Then
            vParent1.StoreRowIndexWithFocus()
            vParent2.StoreRowIndexWithFocus()
            Me.B_app_colorTableAdapter.Fill(Me.TheDataSet.b_app_color)
            Me.B_app_parameterTableAdapter.Fill(Me.TheDataSet.b_app_parameter)
            vParent2.ResetFocusRow()
            vParent1.ResetFocusRow()
        End If
    End Sub
#End Region
#Region "ListBox"
    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim curItem As String = ListBox1.SelectedItem.ToString()
        Try
            TextBox1.BackColor = Color.FromName(curItem)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ListBox1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBox1.MouseDoubleClick
        Dim strColor = ListBox1.SelectedItem.ToString()
        If Not dgParent2.CurrentRow Is Nothing Then
            dgParent2.CurrentRow.Cells("ValueString").Value = strColor
            dgParent2.CurrentRow.Cells("Color").Style.BackColor = Color.FromName(strColor)
            dgParent2.UpdateData()
        End If
    End Sub
#End Region
#Region "Scroll_Resize"
    'And dont forget the filter boxes on the colours.
    'Private Sub frm_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
    Protected Overrides Sub frm_Layout(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles MyBase.Layout
        If TestActiveMDIChild() = True Then

            If Not dgParent1 Is Nothing And Not dgParent2 Is Nothing Then
                Dim iDelta = 10 + statics.GroupBoxRelativeVerticalLocation

                dgParent1.Height = (Me.Height - 120 - iDelta) / 2
                Dim l As Point
                l = Me.dgParent1.Location
                l.Y = 50
                Me.dgParent1.Location = l
                l = vParent1.gbForFiltersGroupBox.Location
                l.Y = Me.dgParent1.Location.Y - statics.GroupBoxRelativeVerticalLocation
                Me.vParent1.gbForFiltersGroupBox.Location = l

                dgParent2.Height = (Me.Height - 120 - iDelta) / 2
                l = Me.dgParent2.Location
                l.Y = Me.dgParent1.Location.Y + dgParent1.Height + iDelta
                Me.dgParent2.Location = l

                If Not vParent2 Is Nothing Then
                    l = vParent2.gbForFiltersGroupBox.Location
                    l.Y = Me.dgParent2.Location.Y - statics.GroupBoxRelativeVerticalLocation
                    Me.vParent2.gbForFiltersGroupBox.Location = l
                End If

                ListBox1.Height = (Me.Height - 120 - iDelta) / 2
                l = Me.ListBox1.Location
                l.Y = Me.dgParent2.Location.Y
                Me.ListBox1.Location = l

                l = Me.TextBox1.Location
                l.Y = Me.dgParent2.Location.Y
                Me.TextBox1.Location = l
            End If
        End If
    End Sub
#End Region

End Class