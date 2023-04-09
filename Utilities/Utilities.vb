'------------------------------------------------
'Name: Module Utilities.vb.
'Function: A collection of functions/subs.
'Copyright Robin Baines 2006. All rights reserved.
'Created May 2006.
'Notes: 
'Modifications:
'RPB March 2007 Added blnExecuteFunction 
'20100105 RPB modified DisplayAForm by Trimming strHeader and strFormName.
'------------------------------------------------
Imports System
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Security.Principal

Public Class Utilities
    Const strGlobalCDateUndefined = "9999999"
#Region "Form"

    Public Sub New()

    End Sub
    Public Function ShouldDelete(ByVal _RowState As DataRowState) As Boolean

        Return (_RowState <> DataRowState.Detached And _RowState <> DataRowState.Added)
    End Function

    Public Function ShouldInsert(ByVal _RowState As DataRowState, ByVal dg As DataGridView) As Boolean
        Return ((_RowState = DataRowState.Detached And dg.IsCurrentRowDirty) Or _RowState = DataRowState.Added)
    End Function

    'Check the following three because the first could be dropped.
    Public Function CanShowAForm(ByVal pParent As Form, ByVal strFormName As String) _
        As Form

        Dim _frm As Form
        _frm = Nothing

        'Check whether form is already open. If not Show otherwise BringToFront.
        Dim ctl() As Control
        ctl = pParent.Controls.Find(strFormName, True)
        If ctl.Length = 0 Then
        Else
            _frm = CType(ctl(0), Form)
            _frm.BringToFront()
        End If
        Return _frm
    End Function

    Public Function frmCanShowMDIForm(ByVal pParent As Form, ByVal strFormName As String, ByVal blnBringToFront As Boolean) _
        As Form

        'Returns fRet as form if exists otherwise Nothing.
        Dim frms() As Form
        Dim f As Form
        Dim fRet As Form = Nothing
        frms = pParent.MdiChildren()
        For Each f In frms
            If f.Name = strFormName Then
                If blnBringToFront = True Then
                    f.BringToFront()
                    fRet = f
                End If
                Exit For
            End If
        Next
        Return fRet
    End Function

    '20100225 Created this to improve on launching a form.
    Public Function blnBringToFrontIfExists(ByVal pParent As Form, ByVal strFormName As String) As Boolean
        Dim blnExists As Boolean = False
        Dim f As Form = Nothing
        Dim bFrmExists As Boolean = False
        strFormName = strFormName.Trim()
        For Each f In pParent.MdiChildren()
            If f.Name = strFormName Then
                'f.WindowState = FormWindowState.Maximized
                f.BringToFront()
                blnExists = True
                Exit For
            End If
        Next
        Return blnExists
    End Function

    Public Sub DisplayAForm(ByVal pParent As Form, ByVal frm As Form, _
                ByVal strHeader As String, ByVal strFormName As String)
        DisplayAForm(pParent, frm, strHeader, strFormName, FormWindowState.Maximized)
    End Sub

    Public Sub DisplayAForm(ByVal pParent As Form, ByVal frm As Form, _
            ByVal strHeader As String, ByVal strFormName As String, ByVal frmState As FormWindowState)

        'Revised April 2008. Use MDIChildren collection.
        'Check whether form is already open. If not Show otherwise BringToFront.
        'Try added around the calls because of problems with Form type conversion.

        '20100105 RPB modified DisplayAForm by Trimming strHeader and strFormName
        strHeader = strHeader.Trim()
        strFormName = strFormName.Trim()
        If blnBringToFrontIfExists(pParent, strFormName) = True Then

            '20100225 
            frm.Dispose()
        Else
            Try
                ShowAForm(pParent, frm, strHeader, strFormName, frmState)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        End If
    End Sub

    Public Sub ShowAForm(ByVal pParent As Form, ByVal frm As Form, ByVal strHeader As String, _
        ByVal strFormName As String)
        ShowAForm(pParent, frm, strHeader, strFormName, FormWindowState.Maximized)
    End Sub

    Public Sub ShowAForm(ByVal pParent As Form, ByVal frm As Form, ByVal strHeader As String, _
        ByVal strFormName As String, ByVal frmState As FormWindowState)

        'Note that a form with a valid MDI parent can not be added to the Controls collection of the parent.
        Dim cur As Cursor
        cur = pParent.Cursor
        pParent.Cursor = Cursors.WaitCursor
        frm.MdiParent = pParent
        If strHeader.Length <> 0 Then
            frm.Text = strHeader
        End If
        frm.Name = strFormName
        frm.WindowState = frmState
        frm.Show()
        pParent.Cursor = cur
    End Sub

    'Public Sub ShowAForm(ByVal pParent As Form, ByVal frm As Form, ByVal strHeader As String, ByVal strFormName As String)
    '    ShowAForm(pParent, frm, strHeader, strFormName, FormWindowState.Maximized)
    'End Sub
    Public Sub DefineColumn(ByVal tb As DataGridViewColumn, ByVal strFormat As String, _
        ByVal blnBound As Boolean, ByVal strdgName As String, ByVal strName As String, ByVal strHeader As String, _
        ByVal iwidth As Integer, _
        ByVal blnRO As Boolean, ByVal blnVisible As Boolean, ByVal strPrintFilter As String, _
        ByVal sdcColor As System.Drawing.Color)

        DefineColumn(tb, strFormat, blnBound, strdgName, strName, strHeader, iwidth, _
                blnRO, blnVisible, strPrintFilter, sdcColor, False)
    End Sub

    Public Sub DefineColumn(ByVal tb As DataGridViewColumn, ByVal strFormat As String, _
            ByVal blnBound As Boolean, ByVal strdgName As String, ByVal strName As String, ByVal strHeader As String, _
            ByVal iwidth As Integer, _
            ByVal blnRO As Boolean, ByVal blnVisible As Boolean, ByVal strPrintFilter As String)

        DefineColumn(tb, strFormat, blnBound, strdgName, strName, strHeader, iwidth, _
                blnRO, blnVisible, strPrintFilter, False)
    End Sub

    Public Sub DefineColumn(ByVal tb As DataGridViewColumn, ByVal strFormat As String, _
            ByVal blnBound As Boolean, ByVal strdgName As String, ByVal strName As String, ByVal strHeader As String, _
            ByVal iwidth As Integer, _
            ByVal blnRO As Boolean, ByVal blnVisible As Boolean, ByVal strPrintFilter As String, ByVal blnBold As Boolean)
        Dim sdcColor As System.Drawing.Color
        If blnRO Then
            sdcColor = Drawing.Color.FromKnownColor(System.Drawing.KnownColor.Control)
        Else
            sdcColor = Drawing.Color.White
        End If

        DefineColumn(tb, strFormat, blnBound, strdgName, strName, strHeader, iwidth, _
                blnRO, blnVisible, strPrintFilter, sdcColor, blnBold)
    End Sub
    Dim strForm As String = ""
    Dim strTable As String = ""
    Dim iSequence As Integer = 0
    Public Sub SetTableNames(ByVal _strForm As String, ByVal _strTable As String)
        strForm = _strForm
        strTable = _strTable
        iSequence = 0

    End Sub
    Public Sub DefineColumn(ByVal tb As DataGridViewColumn, ByVal strFormat As String, _
        ByVal blnBound As Boolean, ByVal strdgName As String, ByVal strName As String, ByVal strHeader As String, _
        ByVal iwidth As Integer, _
        ByVal blnRO As Boolean, ByVal blnVisible As Boolean, ByVal strPrintFilter As String, _
        ByVal sdcColor As System.Drawing.Color, ByVal blnBold As Boolean)

        If blnBound Then tb.DataPropertyName = strName

        'If blnRO = True Then
        ' sdcColor = Drawing.Color.FromKnownColor(System.Drawing.KnownColor.Control)
        'End If
        If strFormat.Length <> 0 Then
            Dim tbStyle As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
            tbStyle.Format = strFormat

            'Mod RPB March 2007. Set to bold if necessary.
            If blnBold = True Then
                tbStyle.Font = New Font(tb.InheritedStyle.Font, FontStyle.Bold)
            End If
            tbStyle.BackColor = sdcColor
            If strFormat.StartsWith("N") Then
                tbStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End If
            tb.DefaultCellStyle = tbStyle
        Else
            tb.DefaultCellStyle.BackColor = sdcColor
        End If

        If strHeader.Length <> 0 Then
            tb.HeaderText = strHeader
        Else
            tb.HeaderText = strName
        End If
        tb.HeaderCell.Style.Font = New Font("Arial", 11, FontStyle.Bold, GraphicsUnit.Pixel)
        tb.HeaderCell.Style.BackColor = Color.AliceBlue
        'Control.DefaultFont, FontStyle.Bold)
        tb.Name = strdgName + strName
        tb.Width = iwidth
        tb.ReadOnly = blnRO
        tb.Visible = blnVisible
        tb.SortMode = DataGridViewColumnSortMode.Automatic
        tb.Tag = strPrintFilter

        'update from database if available otherwise store.
        If strForm.Length > 0 And strTable.Length > 0 Then
            statics.get_v_form_tble_column(strForm, strTable, strName, _
                        strHeader, strFormat, iwidth, blnVisible, blnVisible, blnBold, iSequence)
            iSequence = iSequence + 10
        End If
       
    End Sub
    Public Sub DefineComboBoxColumn(ByVal tb As DataGridViewComboBoxColumn, ByVal strFormat As String, _
        ByVal blnBound As Boolean, ByVal strdgName As String, ByVal strName As String, ByVal strHeader As String, _
        ByVal iwidth As Integer, _
        ByVal blnRO As Boolean, ByVal blnVisible As Boolean, ByVal strPrintFilter As String, _
        ByVal dsBindingSource As BindingSource, ByVal strMember As String)

        Dim sdcColor As System.Drawing.Color
        If blnBound Then tb.DataPropertyName = strName
        If blnRO Then
            sdcColor = Drawing.Color.FromKnownColor(System.Drawing.KnownColor.Control)
        Else
            sdcColor = Drawing.Color.White
        End If

        If strFormat.Length <> 0 Then
            Dim tbStyle As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
            tbStyle.Format = strFormat
            tbStyle.BackColor = sdcColor
            If strFormat.StartsWith("N") Then
                tbStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            End If
            tb.DefaultCellStyle = tbStyle
        Else
            tb.DefaultCellStyle.BackColor = sdcColor
        End If

        If strHeader.Length <> 0 Then
            tb.HeaderText = strHeader
        Else
            tb.HeaderText = strName
        End If

        tb.HeaderCell.Style.Font = New Font("Arial", 11, FontStyle.Bold, GraphicsUnit.Pixel)
        tb.HeaderCell.Style.BackColor = Color.AliceBlue
        tb.Name = strdgName + strName
        tb.Width = iwidth
        tb.ReadOnly = blnRO
        tb.Visible = blnVisible
        tb.DataSource = dsBindingSource
        tb.DisplayMember = strMember
        tb.ValueMember = strMember
        tb.Tag = strPrintFilter
        tb.SortMode = DataGridViewColumnSortMode.Automatic

    End Sub

#End Region

#Region "Filter"
    Public Function CheckTag(ByVal iColumn As DataGridViewColumn, ByVal strTagFilter As String) As Boolean

        'Return False if the character in the Tag of the Column is contained in the TagFiler string.
        'Otherwise return True.
        If IsNothing(iColumn.Tag) Then Return True
        If iColumn.Tag.ToString.Length = 0 Then
            Return True
        End If
        If strTagFilter.Contains(iColumn.Tag) Then Return False
        Return True
    End Function

    Const tbHeight = 20
    Public Sub CreateAFilterBox(ByRef tb As TextBox, ByVal strField As String, _
    ByRef tb_TextChanged As EventHandler, _
    ByVal FTbs As Dictionary(Of Control, String), ByVal Controls As Control.ControlCollection)

        tb = New System.Windows.Forms.TextBox
        tb.Name = "tb" & strField & "Find"
        AddHandler tb.TextChanged, tb_TextChanged
        FTbs.Add(tb, strField)
        Controls.Add(tb)
    End Sub
    Public Sub CreateAButton(ByRef b As Button, ByVal strField As String, _
        ByRef tb_TextChanged As EventHandler, _
        ByVal FTbs As Dictionary(Of Control, String), ByVal Controls As Control.ControlCollection)

        b = New System.Windows.Forms.Button
        b.Name = "b" & strField & "Find"
        AddHandler b.TextChanged, tb_TextChanged
        FTbs.Add(b, strField)
        Controls.Add(b)
    End Sub

    Public Sub CreateACheckBox(ByRef tb As CheckBox, ByVal strField As String, _
    ByRef tb_TextChanged As EventHandler, _
    ByVal FTbs As Dictionary(Of Control, String), ByVal Controls As Control.ControlCollection)

        tb = New System.Windows.Forms.CheckBox
        tb.ThreeState = True
        tb.CheckState = CheckState.Indeterminate
        tb.Name = "cb" & strField & "Find"

        'AddHandler tb.CheckStateChanged, tb_TextChanged
        AddHandler tb.CheckStateChanged, tb_TextChanged
        FTbs.Add(tb, strField)
        Controls.Add(tb)
    End Sub
    Public Function GetLeftOfColumnInGrid(ByVal dg As DataGridView, ByVal col As String) As Integer

        'Return the sum of the widths of the datagridview columns up to but not including the col parameter.
        Dim i As Integer
        Dim w As Integer

        'RPB Feb 2007. Started adjusting the RowHeader column so needed to use the actual value here.
        w = dg.RowHeadersWidth
        i = 0
        While i < dg.Columns.Count
            If dg.Columns(i).Name = col Then
                Exit While
            End If
            'Debug.Print(dg.Columns(i).Name)
            'Debug.Print(dg.Columns(i).Visible)
            If dg.Columns(i).Visible = True Then
                w = w + dg.Columns(i).Width
            End If
            'Debug.Print(dg.Columns(i).Name)
            'Debug.Print(dg.Columns(i).Visible)
            i = i + 1
        End While
        Return w
    End Function

    Public Sub AdjustFilterBoxes(ByVal FindTbs As Dictionary(Of Control, String), ByVal dg As DataGridView)

        'Place the filter boxes above the columns in the datagridview.
        Dim tab As Integer
        tab = 1

        'If dg is in a groupbox then the top position needs to be corrected.
        'Dim gb As GroupBox
        'Dim iOffset As Integer
        'If dg.Parent.ToString.IndexOf("GroupBox") <> -1 Then
        ' gb = CType(dg.Parent, GroupBox)
        ' iOffset = gb.Top
        ' Else
        ' iOffset = 0
        ' End If

        For Each tbEntry As KeyValuePair(Of Control, String) In FindTbs
            'Debug.Print(dg.Columns(0).Name)
            'Debug.Print(dg.Columns(0).Visible)

            '            tbEntry.Key.Location = New System.Drawing.Point(dg.Left + GetLeftOfColumnInGrid(dg, dg.Name & tbEntry.Value), dg.Top - tbHeight + iOffset)
            tbEntry.Key.Location = New System.Drawing.Point(dg.Left + GetLeftOfColumnInGrid(dg, dg.Name & tbEntry.Value), dg.Top - tbHeight)
            'Debug.Print(dg.Columns(0).Name)
            'Debug.Print(dg.Columns(0).Visible)

            tbEntry.Key.Size = New System.Drawing.Size(dg.Columns(dg.Name & tbEntry.Value).Width, tbHeight)
            'Debug.Print(dg.Columns(0).Name)
            'Debug.Print(dg.Columns(0).Visible)

            tbEntry.Key.TabIndex = tab
            'Debug.Print(dg.Columns(0).Name)
            'Debug.Print(dg.Columns(0).Visible)

            tbEntry.Key.Visible = dg.Columns(dg.Name & tbEntry.Value).Visible
            'Debug.Print(dg.Columns(0).Name)
            'Debug.Print(dg.Columns(0).Visible)

            If tbEntry.Key.ToString.Contains("TextBox") = True Then
                Dim tb As TextBox
                tb = CType(tbEntry.Key, TextBox)
                tb.TextAlign = HorizontalAlignment.Left
                'Debug.Print(dg.Columns(0).Name)
                'Debug.Print(dg.Columns(0).Visible)
            End If

            'THIS ACTIVATES the 1st column in the tabs!!!!
            'tbEntry.Key.Text = "*"
            'Debug.Print(dg.Columns(0).Name)
            'Debug.Print(dg.Columns(0).Visible)

            tab = tab + 1
            'Debug.Print(dg.Columns(0).Name)
            'Debug.Print(dg.Columns(0).Visible)

        Next
    End Sub
    'Public Sub SelectFirstTb(ByVal bs As BindingSource, ByVal FindTbs As Dictionary(Of TextBox, String))
    Public Sub SelectFirstTb(ByVal FindTbs As Dictionary(Of TextBox, String))

        'Give the first filter textbox the focus.
        For Each tbEntry As KeyValuePair(Of TextBox, String) In FindTbs
            If tbEntry.Key.TabIndex = 1 Then
                tbEntry.Key.Select()
                Exit For
            End If
        Next
    End Sub

    Public Sub ResetFilter(ByVal bs As BindingSource, ByVal FindTbs As Dictionary(Of Control, String))

        'Remove all strings from the filter textboxes. Is used when re-showing a form as a dialog.
        For Each tbEntry As KeyValuePair(Of Control, String) In FindTbs
            tbEntry.Key.Text = ""
            'tbEntry.Key.Text = "*"
        Next
        bs.RemoveFilter()
        'MakeFilter(bs, FindTbs)
    End Sub


    Public Sub MakeFilterFromButton(ByVal bs As BindingSource, ByVal FindTbs As Dictionary(Of Control, String), _
        ByVal blnAction As Boolean)

        Dim strF As String
        strF = ""
        Dim btn As Button
        For Each tbEntry As KeyValuePair(Of Control, String) In FindTbs
            If tbEntry.Key.ToString.Contains("Button") = True Then
                btn = CType(tbEntry.Key, Button)
                If btn.Tag.Length <> 0 Then
                    If strF.Length <> 0 Then
                        strF = strF & " and "
                    End If
                    If blnAction = True Then
                        strF = strF & " " & tbEntry.Value & " <> " & btn.Tag.Replace("'", "''")
                    Else
                        strF = strF & " "
                    End If
                End If
            End If
        Next
        Try
            strF = strF.Trim
            If strF.Trim.Length = 0 Then
                bs.RemoveFilter()
            Else
                bs.Filter = strF
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly)
        End Try
    End Sub
    Public Function GetBoundColumnName(ByVal dgColumnName As String, ByVal dg As DataGridView) As String

        'Return the name of the bound column.
        Dim strRet As String = ""
        Try
            strRet = dg.Columns(dgColumnName).DataPropertyName
        Catch ex As Exception
        End Try
        Return strRet
    End Function
    Public Function GetBoundColumnType(ByVal dgColumnName As String, ByVal dg As DataGridView) As String

        'Return the type of the bound column.
        'For example System.String or System.Int32.
        Dim strRet As String = ""
        Try
            strRet = dg.Columns(dgColumnName).ValueType.ToString()
        Catch ex As Exception
        End Try
        Return strRet
    End Function
    Public Sub MakeFilter(ByVal bs As BindingSource, ByVal FindTbs As Dictionary(Of Control, String), ByVal dg As DataGridView)

        'RPB July 2008. A safer version of MakeFilter which needs the datagrid to do the extra checking.
        'Iterate through the dictionary of find list boxes to construct the filter and then set the filter.
        Dim strF As String
        strF = ""
        For Each tbEntry As KeyValuePair(Of Control, String) In FindTbs
            If tbEntry.Key.ToString.Contains("TextBox") = True Then
                If tbEntry.Key.Text.Trim.Length <> 0 Then

                    'RPB July 2008
                    'Get the bound column name of the datagrid.
                    Dim strDataPropertyName = GetBoundColumnName(tbEntry.Value, dg)
                    Dim strDataPropertyType As String = GetBoundColumnType(tbEntry.Value, dg)
                    If strDataPropertyName <> "" Then
                        If strF.Length <> 0 Then
                            strF = strF & " and "
                        End If
                        If strDataPropertyType.StartsWith("System.String") Then
                            strF = strF & " " & strDataPropertyName & " Like '" & tbEntry.Key.Text.Replace("'", "''") & "*' "
                        Else
                            strF = strF & " " & strDataPropertyName & " = " & tbEntry.Key.Text.Replace("'", "''")
                        End If
                    End If
                    'strF = strF & " " & tbEntry.Value & " Like '" & tbEntry.Key.Text.Replace("'", "''") & "' "

                End If
            Else
                If tbEntry.Key.ToString.Contains("CheckBox") = True Then
                    Dim tb As CheckBox
                    tb = CType(tbEntry.Key, CheckBox)
                    If tb.CheckState = CheckState.Checked Then
                        If strF.Length <> 0 Then
                            strF = strF & " and "
                        End If
                        strF = strF & " " & tbEntry.Value & " = 1 "
                    End If
                    If tb.CheckState = CheckState.Unchecked Then
                        If strF.Length <> 0 Then
                            strF = strF & " and "
                        End If
                        strF = strF & " " & tbEntry.Value & " = 0 "
                    End If
                    If tb.CheckState = CheckState.Indeterminate Then
                    End If
                End If
            End If
        Next
        Try
            strF = strF.Trim
            If strF.Trim.Length = 0 Then
                bs.RemoveFilter()
            Else
                bs.Filter = strF
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly)
        End Try

    End Sub
    Public Sub MakeFilter(ByVal bs As BindingSource, ByVal FindTbs As Dictionary(Of Control, String))

        'Iterate through the dictionary of find list boxes to construct the filter and then set the filter.
        Dim strF As String
        strF = ""
        For Each tbEntry As KeyValuePair(Of Control, String) In FindTbs
            If tbEntry.Key.ToString.Contains("TextBox") = True Then
                If tbEntry.Key.Text.Trim.Length <> 0 Then
                    If strF.Length <> 0 Then
                        strF = strF & " and "
                    End If
                    strF = strF & " " & tbEntry.Value & " Like '" & tbEntry.Key.Text.Replace("'", "''") & "*' "
                    'strF = strF & " " & tbEntry.Value & " Like '" & tbEntry.Key.Text.Replace("'", "''") & "' "
                End If
            Else
                If tbEntry.Key.ToString.Contains("CheckBox") = True Then
                    Dim tb As CheckBox
                    tb = CType(tbEntry.Key, CheckBox)
                    If tb.CheckState = CheckState.Checked Then
                        If strF.Length <> 0 Then
                            strF = strF & " and "
                        End If
                        strF = strF & " " & tbEntry.Value & " = 1 "
                    End If
                    If tb.CheckState = CheckState.Unchecked Then
                        If strF.Length <> 0 Then
                            strF = strF & " and "
                        End If
                        strF = strF & " " & tbEntry.Value & " = 0 "
                    End If
                    If tb.CheckState = CheckState.Indeterminate Then
                    End If
                End If
            End If
        Next
        Try
            strF = strF.Trim
            If strF.Trim.Length = 0 Then
                bs.RemoveFilter()
            Else
                bs.Filter = strF
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly)
        End Try

    End Sub
    Public Function ColumnDoubleClick(ByVal bs As BindingSource, _
        ByVal FTbs As Dictionary(Of Control, String), _
        ByVal sender As System.Object, _
        ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) As Boolean
        Dim blnRet = False
        If e.RowIndex >= 0 And e.ColumnIndex >= 0 Then
            Dim dg As DataGridView
            dg = CType(sender, DataGridView)
            Dim strColumnName As String = dg.Columns(e.ColumnIndex).Name.Remove(0, dg.Name.Length)
            ColumnDoubleClick(bs, FTbs, strColumnName, dg.Rows(e.RowIndex).Cells(e.ColumnIndex).Value())
        End If
        'Dim blnRet = False
        'If e.RowIndex >= 0 And e.ColumnIndex >= 0 Then
        '    Dim dg As DataGridView
        '    dg = CType(sender, DataGridView)
        '    Debug.Print(dg.Columns(e.ColumnIndex).Name)
        '    For Each tbEntry As KeyValuePair(Of Control, String) In FTbs
        '        Dim strLB As String
        '        Dim strDG As String
        '        strLB = tbEntry.Key.Name.Remove(0, 2)       'tb
        '        strLB = strLB.Remove(strLB.LastIndexOf("Find"), 4)
        '        strDG = dg.Columns(e.ColumnIndex).Name.Remove(0, dg.Name.Length)
        '        If strLB = strDG Then
        '            tbEntry.Key.Text = dg.Rows(e.RowIndex).Cells(e.ColumnIndex).Value()
        '            MakeFilter(bs, FTbs)
        '            blnRet = True
        '            Exit For
        '        End If
        '    Next
        'End If
    End Function
    Public Function ColumnDoubleClick(ByVal bs As BindingSource, _
    ByVal FTbs As Dictionary(Of Control, String), _
    ByVal strColumnName As String, _
    ByVal strFilterText As String) As Boolean

        'RPB Feb 2008. Lookup the column and place the text in it.
        Dim blnRet = False
        For Each tbEntry As KeyValuePair(Of Control, String) In FTbs
            Dim strLB As String
            strLB = tbEntry.Key.Name.Remove(0, 2)       'tb
            strLB = strLB.Remove(strLB.LastIndexOf("Find"), 4)
            If strLB = strColumnName Then
                tbEntry.Key.Text = strFilterText
                MakeFilter(bs, FTbs)
                blnRet = True
                Exit For
            End If
        Next

    End Function
#End Region
#Region "FormEventHandlers"
    Public Sub DataErrorHandler(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs)
        'This is the standard error handler for data grid columns. Is triggered, for example, 
        'if an non number is entered in an N2 column. 
        'Have not discovered how to define the formatting more accurately. For example Brix 
        'should not only be N2 but should not be negative. This extra checking has been implemented in the 
        'CellValidating handler below. 

        'CHECK: Is needed here to show the user the error text for non valid fields after editing has started.
        'It could be better to let the database throw the exception directly but see comment on 
        'EndEdit in RowValidating.
        Dim dg As DataGridView = CType(sender, DataGridView)
        Dim rv As DataGridViewCell
        rv = dg.CurrentCell
        e.ThrowException = False
        'Debug.Print("data error")
        e.Cancel = True
        dg.CurrentRow.ErrorText = "Data not saved." & vbCrLf & e.Exception.Message
    End Sub

    Public Function Handle_DataGridView_RowValidating(ByVal EventName As String, ByRef TableAdapter As System.Object, ByVal sender As System.Object, _
    ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs, ByVal iCheckColumn As Integer) As Boolean
        'Return true if Update.
        Dim blnRet As Boolean
        blnRet = False
        Dim dg As DataGridView = CType(sender, DataGridView)
        Try

            Dim bs As BindingSource = CType(dg.DataSource, BindingSource)
            Dim aRow As System.Data.DataRow
            Dim sD As System.Data.DataRowView
            sD = bs.Current

            If Not sD Is Nothing Then
                aRow = sD.Row
                Select Case aRow.RowState
                    Case DataRowState.Added
                    Case DataRowState.Deleted
                    Case DataRowState.Detached
                        If Not aRow.IsNull(iCheckColumn) Then
                            Try 'this can fail is another column in the row may not be null but is.
                                'just ignore.
                                bs.EndEdit()
                            Catch ex As Exception
                                'MsgBox(ex.Message)
                                'e.Cancel = True
                            End Try
                        End If
                    Case DataRowState.Modified
                    Case DataRowState.Unchanged
                        bs.EndEdit()
                End Select
                'Debug.Print("--->" & aRow.RowState.ToString())
                Select Case aRow.RowState
                    Case DataRowState.Added
                        CallByName(TableAdapter, "Update", CallType.Method, aRow)
                        blnRet = True
                    Case DataRowState.Deleted
                    Case DataRowState.Detached
                    Case DataRowState.Modified
                        CallByName(TableAdapter, "Update", CallType.Method, aRow)
                        blnRet = True
                    Case DataRowState.Unchanged
                End Select
            End If
        Catch ex As Exception
            'RPB Jan 2008 Gives a problem when using proposed data on a ComboBox input field.
            'If ex.Message.IndexOf("no Proposed data") = 0 Then
            MsgBox("EXCEPTION: " & ex.Message & " ")
            e.Cancel = True
            'Else
            'dg.CurrentRow.ErrorText = ""
            'End If
        Finally
            'Debug.Print("--->" & dr.RowState.ToString())
        End Try
        Return blnRet
    End Function

    Public Sub Handle_DataGridView_UserDeletingRow(ByVal EventName As String, ByRef TableAdapter As System.Object, ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs)
        Dim dg As DataGridView = CType(sender, DataGridView)
        Dim bs As BindingSource = CType(dg.DataSource, BindingSource)
        Try
            Debug.Print(EventName & " " & e.Row.Cells(0).Value)
            CallByName(TableAdapter, "Delete", CallType.Method, e.Row.Cells(0).Value)
        Catch ex As Exception
            MsgBox("EXCEPTION: " & ex.Message & " ")
            e.Cancel = True
        End Try
    End Sub

    Public Sub Handle_DataGridView_UserDeletingRow2(ByVal EventName As String, ByRef TableAdapter As System.Object, ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs)
        Dim dg As DataGridView = CType(sender, DataGridView)
        Dim bs As BindingSource = CType(dg.DataSource, BindingSource)
        Try
            Debug.Print(EventName & " " & e.Row.Cells(0).Value)
            CallByName(TableAdapter, "Delete", CallType.Method, e.Row.Cells(0).Value, e.Row.Cells(1).Value)
        Catch ex As Exception
            MsgBox("EXCEPTION: " & ex.Message & " ")
            e.Cancel = True
        End Try
    End Sub

    Public Sub Handle_DataGridView_UserDeletingRow3(ByVal EventName As String, ByRef TableAdapter As System.Object, ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowCancelEventArgs)
        Dim dg As DataGridView = CType(sender, DataGridView)
        Dim bs As BindingSource = CType(dg.DataSource, BindingSource)
        Try
            Debug.Print(EventName & " " & e.Row.Cells(0).Value)
            CallByName(TableAdapter, "Delete", CallType.Method, e.Row.Cells(0).Value, e.Row.Cells(1).Value, e.Row.Cells(2).Value)
        Catch ex As Exception
            MsgBox("EXCEPTION: " & ex.Message & " ")
            e.Cancel = True
        End Try
    End Sub

#End Region

#Region "Database"
    Public Function GetDataSetODBC(ByVal cn As Odbc.OdbcConnection, ByVal sqlQuery As String) As DataSet

        'Execute a select statement and load into a dataset.
        Dim DataSet As New DataSet
        Dim command As New Odbc.OdbcCommand
        command = cn.CreateCommand()
        command.CommandText = sqlQuery
        Dim dataAdapter As Odbc.OdbcDataAdapter
        dataAdapter = New Odbc.OdbcDataAdapter(command)
        dataAdapter.Fill(DataSet)
        'cn.Close()
        Return DataSet
    End Function

    Public Function GetDataSet(ByVal cn As SqlConnection, ByVal sqlQuery As String) As DataSet

        'Execute a select statement and load into a dataset.
        Dim DataSet As New DataSet
        Dim command As New SqlCommand
        command = cn.CreateCommand()
        command.CommandText = sqlQuery
        Dim dataAdapter As SqlDataAdapter
        dataAdapter = New SqlDataAdapter(command)
        dataAdapter.Fill(DataSet)
        'cn.Close()
        Return DataSet
    End Function

    Public Function ExecuteScalar(ByVal cn As SqlConnection, ByVal sqlQuery As String) As Boolean

        Dim strError As String = ""
        'Execute an SQL expression which does not return data (update, insert or delete)
        'RPB April 2008 Added ExecuteScalar with explicit timeout parameter.
        Return ExecuteScalar(strError, cn, sqlQuery, 60)
    End Function

    Public Function ExecuteScalar(ByRef strError As String, ByVal cn As SqlConnection, ByVal sqlQuery As String) As Boolean

        'Execute an SQL expression which does not return data (update, insert or delete)
        'RPB April 2008 Added ExecuteScalar with explicit timeout parameter.
        Return ExecuteScalar(strError, cn, sqlQuery, 60)
    End Function

    Public Function ExecuteScalar(ByRef strError As String, ByVal cn As SqlConnection _
        , ByVal sqlQuery As String, ByVal iTimeout As Integer) As Boolean

        'Execute an SQL expression which does not return data (update, insert or delete)
        Dim blnRet As Boolean
        Dim command As SqlCommand
        strError = ""
        Try
            blnRet = True
            sqlQuery = sqlQuery
            command = New SqlCommand(sqlQuery, cn)
            command.CommandTimeout = iTimeout
            Dim result As Integer
            result = Convert.ToInt32(command.ExecuteScalar())
        Catch ex As Exception

            'Use return string for the error instead of MsgBox because MsgBox fails when used in a service.
            strError = "ExecuteScalar " + ex.Message
            blnRet = False
        End Try
        Return blnRet
    End Function

    Public Function blnExecuteStoredProcedure(ByRef strReturn() As String, _
        ByRef strError As String, ByVal strConnection As String, ByVal strProcedure As String, _
        ByVal iTimeout As Integer) As Boolean

        Dim cn As New SqlConnection
        Dim blnRet As Boolean

        cn.ConnectionString = strConnection
        cn.Open()
        blnRet = blnExecuteStoredProcedure(strReturn, strError, cn, strProcedure, iTimeout)
        cn.Close()
        Return blnRet
    End Function

    Public Function blnExecuteStoredProcedure(ByRef strError As String, _
    ByVal cn As SqlConnection, ByVal strProcedure As String, ByVal iTimeout As Integer) As Boolean
        Dim strReturn(20) As String
        Return blnExecuteStoredProcedure(strReturn, strError, cn, strProcedure, iTimeout)
    End Function

    Public Function blnExecuteStoredProcedure(ByRef strReturn() As String, _
    ByRef strError As String, ByVal cn As SqlConnection, ByVal strProcedure As String, ByVal iTimeout As Integer) As Boolean

        'Execute a stored procedure.
        Dim cmd As System.Data.SqlClient.SqlCommand
        Dim blnRet As Boolean
        Dim reader As System.Data.SqlClient.SqlDataReader
        blnRet = True
        strError = ""
        Try
            cmd = New System.Data.SqlClient.SqlCommand(strProcedure, cn)
            If iTimeout > 30 Then
                cmd.CommandTimeout = iTimeout
            End If
            cmd.CommandType = CommandType.StoredProcedure
            Dim retParam As System.Data.SqlClient.SqlParameter
            retParam = cmd.Parameters.Add("@ReturnValue", SqlDbType.Int)
            retParam.Direction = ParameterDirection.ReturnValue
            reader = cmd.ExecuteReader()
            Dim i As Integer
            i = 0
            While reader.Read() And i < strReturn.Length
                If reader.IsDBNull(0) = False Then
                    strReturn(i) = reader.GetString(0)
                    i = i + 1
                End If
            End While
            reader.Close()

            'retParam is only set after the reader is closed!
            If (retParam.Value <> 1) Then
                blnRet = False

                'Use return string for the error instead of MsgBox because MsgBox fails when used in a service.
                strError = strProcedure & " returned an error."
            End If
        Catch ex As Exception
            blnRet = False
            strError = "Problem in " & strProcedure & ": " & ex.Message
        End Try
        Return blnRet
    End Function

    Public Function blnExecuteStoredProcedure(ByVal cn As SqlConnection, ByVal strProcedure As String, _
        ByVal strParameter As String, ByVal iTimeout As Integer) As Boolean
        Dim strError As String = ""
        Return blnExecuteStoredProcedure(strError, cn, strProcedure, strParameter, iTimeout)
    End Function

    'RPB April 2008 Could simplify by calling function with parametername in the parameter list. See next function.
    Public Function blnExecuteStoredProcedure(ByRef strError As String, _
        ByVal cn As SqlConnection, ByVal strProcedure As String, _
        ByVal strParameter As String, ByVal iTimeout As Integer) As Boolean

        'Execute a stored procedure.
        Dim cmd As System.Data.SqlClient.SqlCommand
        Dim blnRet As Boolean
        Dim reader As System.Data.SqlClient.SqlDataReader
        strError = ""
        blnRet = True
        Try
            cmd = New System.Data.SqlClient.SqlCommand(strProcedure, cn)
            If iTimeout > 30 Then
                cmd.CommandTimeout = iTimeout
            End If
            cmd.CommandType = CommandType.StoredProcedure
            Dim retParam1 As New System.Data.SqlClient.SqlParameter
            retParam1 = cmd.Parameters.Add("@FileName", SqlDbType.NVarChar, 240)
            retParam1.Value = strParameter
            retParam1.Direction = ParameterDirection.Input

            reader = cmd.ExecuteReader()
            Dim i As Integer
            i = 0
            reader.Close()

        Catch ex As Exception
            blnRet = False

            'Use return string for the error instead of MsgBox because MsgBox fails when used in a service.
            strError = "Problem in " & strProcedure & ": " & ex.Message
        End Try
        Return blnRet
    End Function

    Public Function blnExecuteStoredProcedure(ByRef strError As String, _
            ByVal cn As SqlConnection, ByVal strProcedure As String, _
            ByVal strParameter As String, ByVal strParameterName As String, ByVal iTimeout As Integer) As Boolean

        'Execute a stored procedure.
        Dim cmd As System.Data.SqlClient.SqlCommand
        Dim blnRet As Boolean
        Dim reader As System.Data.SqlClient.SqlDataReader
        strError = ""
        blnRet = True
        Try
            cmd = New System.Data.SqlClient.SqlCommand(strProcedure, cn)
            If iTimeout > 30 Then
                cmd.CommandTimeout = iTimeout
            End If
            cmd.CommandType = CommandType.StoredProcedure
            Dim retParam1 As New System.Data.SqlClient.SqlParameter
            retParam1 = cmd.Parameters.Add(strParameterName, SqlDbType.NVarChar, 240)
            retParam1.Value = strParameter
            retParam1.Direction = ParameterDirection.Input

            reader = cmd.ExecuteReader()
            Dim i As Integer
            i = 0
            reader.Close()

        Catch ex As Exception
            blnRet = False

            'Use return string for the error instead of MsgBox because MsgBox fails when used in a service.
            strError = "Problem in " & strProcedure & ": " & ex.Message
        End Try
        Return blnRet
    End Function

    Public Function blnExecuteFunction(ByRef strError As String, _
        ByRef strReturn As String, _
        ByVal cn As SqlConnection, ByVal strFunction As String, ByVal iTimeout As Integer) As Boolean

        Dim blnRet As Boolean
        blnRet = True
        strError = ""
        strReturn = ""
        Dim command As SqlCommand
        Try
            command = New SqlCommand(strFunction, cn)
            command.CommandTimeout = 60
            strReturn = command.ExecuteScalar()
        Catch ex As Exception

            'Use return string for the error instead of MsgBox because MsgBox fails when used in a service.
            strError = "blnExecuteFunction " + ex.Message
            blnRet = False
        End Try
        Return blnRet
    End Function

    Public Function strGetParameter(ByVal cn As SqlConnection, ByVal strParameterId As String) As Integer
        'Use a stored procedure to check whether the order number is present.
        Dim cmd As System.Data.SqlClient.SqlCommand
        Dim iValue As Int32
        Dim iType As Int32
        Dim reader As System.Data.SqlClient.SqlDataReader

        cmd = New System.Data.SqlClient.SqlCommand("fn_GetParameter", cn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add("@Name", SqlDbType.Char)
        cmd.Parameters("@Name").Value = strParameterId

        Dim retParam As System.Data.SqlClient.SqlParameter
        retParam = cmd.Parameters.Add("@ReturnValue", SqlDbType.Int)
        retParam.Direction = ParameterDirection.ReturnValue
        reader = cmd.ExecuteReader()
        If reader.HasRows Then
            If (reader.Read() = True) Then
                If reader.IsDBNull(4) = False Then iType = reader.GetInt32(4)
                If iType = 1 Then
                    If reader.IsDBNull(2) = False Then iValue = reader.GetInt32(2)
                End If
            End If
        End If
        reader.Close()
        Return iValue
    End Function

#End Region

#Region "Security"

    Public Function blnCheckLevel(ByVal ds As DataSet, ByVal strForm As String, ByRef blnRO As Boolean) As Boolean
        Dim blnRet = False

        '20100105 RPB modified blnCheckLevel by Trimming strForm
        strForm = strForm.Trim()
        blnRO = False
        If ds.Tables.Count > 0 Then
            Dim tFormLevels = ds.Tables(0)
            For Each r As DataRow In tFormLevels.Rows
                If r.Item("Form").ToString().ToUpper = "ALL" Then
                    blnRet = True
                    blnRO = r.Item("RO")
                    Exit For
                End If
                If r.Item("Form").ToString().ToUpper = strForm.Trim.ToUpper Then
                    blnRO = r.Item("RO")
                    blnRet = True
                    Exit For
                End If
            Next
        End If
        Return blnRet
    End Function

    Public Function GetSecurityLevel(ByRef ds As DataSet, ByVal strConnection As String) As String

        'Lookup the user in the security table and return associated security value.
        'If user is not found return the lowest level.

        ' Retrieve the Windows account token for the current user.
        Dim strWindowsName As String = ""
        Dim strUser = strGetUserName(strWindowsName)
        If strConnection.ToLower.Contains("dsn=") Then
            Dim cn As New Odbc.OdbcConnection
            cn.ConnectionString = strConnection
            cn.Open()
            ds = GetDataSetODBC(cn, "select * from bUsers u, bFormLevels l " & _
                " where User='" & strUser & "'" & _
                " and u.Level=l.Level")
            cn.Close()
        Else
            Dim cn As New SqlConnection

            cn.ConnectionString = strConnection
            cn.Open()
            ds = GetDataSet(cn, "select * from bUsers u, bFormLevels l " & _
                " where [User]='" & strUser & "'" & _
                " and u.Level=l.Level")
            cn.Close()
        End If

        Return strWindowsName
    End Function

    Private Function LogonUser() As IntPtr
        Dim accountToken As IntPtr = WindowsIdentity.GetCurrent().Token
        Return accountToken
    End Function

    Private Function strGetUserName(ByRef strWindowsName As String)

        Dim logonToken As IntPtr = LogonUser()
        Dim windowsIdentity As New WindowsIdentity(logonToken)
        Dim strUser As String
        strWindowsName = windowsIdentity.Name
        strUser = windowsIdentity.Name
        Dim i As Integer
        i = windowsIdentity.Name.IndexOf("\")
        If i > 0 Then
            strUser = windowsIdentity.Name.Substring(i + 1)
        End If
        Return strUser
    End Function

#End Region

#Region "Dates"
    'use moduledates: namespace dates instead.
    Public Function strGetCJSDate() As String
        Return dates.strGetCJSDate()
    End Function

    Public Function strGetCJSDate(ByVal dt As DateTime) As String
        Return dates.strGetCJSDate(dt)
    End Function

    Public Function strConvertFromCDateTime(ByVal strCDate As String, ByVal strCTime As String) As String
        Return dates.strConvertFromCDateTime(strCDate, strCTime)
    End Function

    Public Function strConvertToTPDateTime(ByVal strAccessDate As String) As String
        Return dates.strConvertToTPDateTime(strAccessDate)
    End Function

    Public Function strGetThisMonth() As String
        Return dates.strGetThisMonth()
    End Function

    Public Function IsCDate(ByVal strDate As String) As Boolean
        Return dates.IsCDate(strDate)
    End Function

    Public Function strIncrementDate(ByVal strDate As String) As String
        Return dates.strIncrementDate(strDate)
    End Function
#End Region

End Class
