'------------------------------------------------
'Name: Module frmManage.vb
'Function: The Security form a.k.a. Management Form.
'Copyright Robin Baines 2008. All rights reserved.
'Created Feb 2008.
'Purpose: 
'Notes: 
'Modifications: 
'20151107 do not allow filters because the limited selection gets saved and that is probably not the idea.
'20151115 comment change while testing maint branch in git.
'20200219 Do not allow dialog DialogSelectForms to be shown if the Security form is Read only.
'------------------------------------------------
Imports System.Windows.Forms
Imports System.Drawing

Public Class frmManage
    Dim vLang As TheDataSet_m_lang
    Dim vtble As TheDataSet_m_tble
    Dim vform_tble As TheDataSet_m_form_tble
    Public vtble_column As TheDataSet_m_tble_column
    Dim vform_tble_column__visibility As TheDataSet_M_form_tble_column__visibility
    Dim vform As TheDataSet_m_form
    Dim vusr As TheDataSet_m_usr
    Dim vform_grp As TheDataSet_m_form_grp
    Dim vform_grp2 As TheDataSet_m_form_grp
    Dim m_form_grp_groupbox As TheDataSet_m_form_grp_groupbox
    Dim vgrp As TheDataSet_m_grp
    Dim vformat As TheDataSet_m_format
    Dim vtble_column_header As TheDataSet_v_tble_column_header
    Dim vusr_log As TheDataSet_v_usr_log
    Dim vtxt As TheDataSet_m_txt
    Dim vtxt_header As TheDataSet_v_txt_header
    Friend WithEvents tsbUnBlock As System.Windows.Forms.ToolStripButton
    Dim blnAll As Boolean
    Dim strForm As String = ""

    Dim m_form_grp_log As TheDataSet_m_form_grp_log
    Dim m_form_grp_log2 As TheDataSet_m_form_grp_log
    Dim m_usr_change_log As TheDataSet_m_usr_change_log

    Dim m_form_grp_log_all As TheDataSet_m_form_grp_log
    Dim m_usr_change_log_all As TheDataSet_m_usr_change_log

#Region "New"
    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(ByVal tsb As ToolStripItem _
               , ByVal strSecurityName As String, ByVal _MainDefs As MainDefinitions)

        MyBase.New(tsb, strSecurityName, _MainDefs)
        InitializeComponent()
        blnAll = True
        strForm = ""
        InitGrids(strSecurityName)
    End Sub

    Public Sub New(ByVal strSecurityName As String, ByVal _MainDefs As MainDefinitions, ByVal _strForm As String)
        MyBase.New(Nothing, strSecurityName, _MainDefs)
        InitializeComponent()
        blnAll = False
        strForm = _strForm
        InitGrids(strSecurityName)
    End Sub

    Private Sub InitGrids(ByVal strSecurityName As String)

        'tables
        vtble = New TheDataSet_m_tble(strSecurityName, Me.M_tbleBindingSource, M_tbleDataGridView, Me.M_tbleTableAdapter, _
                          Me.TheDataSet, _
                          Me.components, _
                          MainDefs, True, Me.tpTables.Controls, Me, True)
        M_tbleDataGridView.AllowUserToDeleteRows = True

        vtble_column = New TheDataSet_m_tble_column(strSecurityName, Me.M_tble_columnBindingSource, M_tble_columnDataGridView, Me.M_tble_columnTableAdapter, _
                          Me.TheDataSet, _
                          Me.components, _
                          MainDefs, blnRO, Me.tpTables.Controls, Me, True)
        M_tble_columnDataGridView.AllowUserToAddRows = False

        vtble_column_header = New TheDataSet_v_tble_column_header(strSecurityName, Me.V_tble_column_headerBindingSource, v_tble_column_headerDataGridView, Me.V_tble_column_headerTableAdapter, _
              Me.TheDataSet, _
              Me.components, _
              MainDefs, blnRO, tpTables.Controls, Me, False)
        v_tble_column_headerDataGridView.AllowUserToAddRows = False


        'forms
        vform_tble = New TheDataSet_m_form_tble(strSecurityName, Me.M_form_tbleBindingSource, M_form_tbleDataGridView, Me.M_form_tbleTableAdapter, _
                        Me.TheDataSet, _
                        Me.components, _
                        MainDefs, True, Me.tpForms.Controls, Me, True)

        M_form_tbleDataGridView.AllowUserToDeleteRows = True

        vform_tble_column__visibility = New TheDataSet_m_form_tble_column__visibility(strSecurityName, Me.M_form_tble_column__visibilityBindingSource, M_form_tble_column__visibilityDataGridView, Me.M_form_tble_column__visibilityTableAdapter, _
                          Me.TheDataSet, _
                          Me.components, _
                          MainDefs, blnRO, Me.tpForms.Controls, Me, True)
        M_form_tble_column__visibilityDataGridView.AllowUserToAddRows = False

        vform = New TheDataSet_m_form(strSecurityName, Me.M_formBindingSource, M_formDataGridView, Me.M_formTableAdapter, _
                   Me.TheDataSet, _
                   Me.components, _
                   MainDefs, True, Me.tpForms.Controls, Me, True)
        M_formDataGridView.AllowUserToDeleteRows = True


        '20141229 show in which groups a form is defined.
        'Note that the table name is altered to allow independent adjustment of the columns.
        vform_grp2 = New TheDataSet_m_form_grp(strSecurityName, Me.M_form_grpBindingSource, Me.dgGrp, Me.M_form_grpTableAdapter, _
                  Me.TheDataSet, _
                  Me.components, _
                  MainDefs, True, Me.tpForms.Controls, Me, False, "m_frm_grp")

        m_form_grp_log = New TheDataSet_m_form_grp_log(strSecurityName, Me.M_form_grp_logBindingSource, Me.dgform_grp_log, Me.M_form_grp_logTableAdapter, _
                  Me.TheDataSet, _
                  Me.components, _
                  MainDefs, True, Me.tpForms.Controls, Me, False, "m_form_grp_log")

        m_form_grp_log2 = New TheDataSet_m_form_grp_log(strSecurityName, Me.M_form_grp_logBindingSource, Me.dgform_grp_log2, Me.M_form_grp_logTableAdapter, _
                  Me.TheDataSet, _
                  Me.components, _
                  MainDefs, True, Me.tpGroup.Controls, Me, True, "m_form_grp_log2")

        If blnAll = True Then
            vLang = New TheDataSet_m_lang(strSecurityName, Me.M_langBindingSource, M_langDataGridView, Me.M_langTableAdapter, _
                            Me.TheDataSet, _
                            Me.components, _
                            MainDefs, blnRO, Me.tpLanguage.Controls, Me, False)

            vusr = New TheDataSet_m_usr(strSecurityName, Me.M_usrBindingSource, M_usrDataGridView, Me.M_usrTableAdapter, _
                      Me.TheDataSet, _
                      Me.components, _
                      MainDefs, blnRO, Me.tpUsers.Controls, Me, True)

            vusr_log = New TheDataSet_v_usr_log(strSecurityName, Me.V_usr_logBindingSource, dgChild1, Me.V_usr_logTableAdapter, _
                              Me.TheDataSet, _
                              Me.components, _
                              MainDefs, True, Me.tpUsers.Controls, Me, False)

            '20141230
            m_usr_change_log = New TheDataSet_m_usr_change_log(strSecurityName, Me.M_usr_change_logBindingSource, dg_usr_change_log, Me.M_usr_change_logTableAdapter, _
                              Me.TheDataSet, _
                              Me.components, _
                              MainDefs, True, Me.tpUsers.Controls, Me, False, "m_usr_change_log")

            vform_grp = New TheDataSet_m_form_grp(strSecurityName, Me.M_form_grpBindingSource, M_form_grpDataGridView, Me.M_form_grpTableAdapter, _
                      Me.TheDataSet, _
                      Me.components, _
                       MainDefs, True, Me.tpGroup.Controls, Me, False, "m_form_grp")
            '20151107 do not allow filters because the limited selection gets saved and that is probably not the idea.
            'MainDefs, True, Me.tpGroup.Controls, Me, True, "m_form_grp")

            m_form_grp_groupbox = New TheDataSet_m_form_grp_groupbox(strSecurityName, Me.M_form_grp_groupboxBindingSource, _
                     M_form_grp_groupboxdgvEnter, Me.M_form_grp_groupboxTableAdapter, _
                     Me.TheDataSet, _
                     Me.components, _
                     MainDefs, blnRO, Me.tpGroup.Controls, Me, True)
            M_form_grp_groupboxdgvEnter.AllowUserToAddRows = False

            vgrp = New TheDataSet_m_grp(strSecurityName, Me.M_grpBindingSource, M_grpDataGridView, Me.M_grpTableAdapter, _
                      Me.TheDataSet, _
                      Me.components, _
                      MainDefs, blnRO, Me.tpGroup.Controls, Me, True)

            vformat = New TheDataSet_m_format(strSecurityName, Me.M_formatBindingSource, M_formatDataGridView, Me.M_formatTableAdapter, _
                      Me.TheDataSet, _
                      Me.components, _
                      MainDefs, True, Controls, Me, False)

            vtxt = New TheDataSet_m_txt(strSecurityName, Me.M_txtBindingSource, M_txtDataGridView, Me.M_txtTableAdapter, _
                    Me.TheDataSet, _
                    Me.components, _
                    MainDefs, True, tpTexts.Controls, Me, True)
            M_txtDataGridView.AllowUserToDeleteRows = True

            vtxt_header = New TheDataSet_v_txt_header(strSecurityName, Me.V_txt_headerBindingSource, v_txt_headerDataGridView, Me.V_txt_headerTableAdapter, _
                    Me.TheDataSet, _
                    Me.components, _
                    MainDefs, blnRO, tpTexts.Controls, Me, False)
            v_txt_headerDataGridView.AllowUserToAddRows = False

            m_form_grp_log_all = New TheDataSet_m_form_grp_log(strSecurityName, Me.M_form_grp_logBindingSource, Me.dg_form_grp_log, Me.M_form_grp_logTableAdapter, _
                Me.TheDataSet, _
                Me.components, _
                MainDefs, True, Me.tpGrpFormLog.Controls, Me, True, "m_usr_change_log_all")

            m_usr_change_log_all = New TheDataSet_m_usr_change_log(strSecurityName, Me.M_usr_change_logBindingSource, Me.dg_m_usr_change_log, Me.M_usr_change_logTableAdapter, _
                              Me.TheDataSet, _
                              Me.components, _
                              MainDefs, True, Me.tpUserChangeLog.Controls, Me, True, "m_form_grp_log_all")

        Else
            TabControl1.TabPages.Remove(tpLanguage)
            TabControl1.TabPages.Remove(tpUsers)
            TabControl1.TabPages.Remove(tpGroup)
            TabControl1.TabPages.Remove(tpFormat)
            TabControl1.TabPages.Remove(tpTexts)
            TabControl1.TabPages.Remove(tpGrpFormLog)
            TabControl1.TabPages.Remove(tpUserChangeLog)
        End If

        iInitialFormHeight = 1022 + 60

        Me.tsbUnBlock = Me.CreateTsb("tsbUnBlock", "UnBlock User", True, True)
        tsbUnBlock.Visible = False
        Me.SwitchOffUpdate()    'this is too difficult to couple up in the tabs.
        Me.SwitchOffPrintDetail()
    End Sub
#End Region
#Region "Load"

    Protected Sub AdjustHeightsDataGridViews()


        For Each dgc As dgColumns In vGrids
            If Not dgc.dg.Name = "v_txt_headerDataGridView" And Not dgc.dg.Name = "M_form_grp_groupboxdgvEnter" _
                                       And Not dgc.dg.Name = "M_form_tbleDataGridView" _
                                       And Not dgc.dg.Name = "M_form_tble_column__visibilityDataGridView" Then
                dgc.SetHeight(Me.TabControl1.Height - 69) ' - dgc.dg.Location.Y)
            End If
        Next
        Return
    End Sub

    ''' <summary>
    ''' The Users tab shows the User statistics in textboxes instead of a DataGridView.
    ''' This Sub builds the label/textboxes based on whether the associated DataGridView column is visible.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CreateuserStatisticsBoxes(ByVal iParentWidth As Integer)
        Dim ALabel As System.Windows.Forms.Label
        Dim ATextBox As System.Windows.Forms.TextBox
        Dim ACheckBox As System.Windows.Forms.CheckBox
        Dim iY = 20
        Dim iX = iParentWidth + 100
        Dim iCount As Integer = 0

        EmailTextBox.Location = New System.Drawing.Point(iX + 300, iY)
        Dim EmailLabel As System.Windows.Forms.Label
        EmailLabel = New System.Windows.Forms.Label
        EmailLabel.AutoSize = True
        EmailLabel.Location = New System.Drawing.Point(iX, iY)
        EmailLabel.Name = "EmailLabel"
        EmailLabel.Size = New System.Drawing.Size(53, 13)
        EmailLabel.TabIndex = 2 + iCount
        EmailLabel.Text = "Email"
        Me.tpUsers.Controls.Add(EmailLabel)
        iY = iY + 30
        TelephoneTextBox.Location = New System.Drawing.Point(iX + 300, iY)
        Dim TelephoneLabel As System.Windows.Forms.Label
        TelephoneLabel = New System.Windows.Forms.Label
        TelephoneLabel.AutoSize = True
        TelephoneLabel.Location = New System.Drawing.Point(iX, iY)
        TelephoneLabel.Name = "TelephoneLabel"
        TelephoneLabel.Size = New System.Drawing.Size(53, 13)
        TelephoneLabel.TabIndex = 2 + iCount
        TelephoneLabel.Text = "Telephone"
        Me.tpUsers.Controls.Add(TelephoneLabel)
        iY = iY + 30
        For Each col As DataGridViewColumn In dgChild1.Columns
            If col.Visible = True Then
                ALabel = New System.Windows.Forms.Label
                ALabel.AutoSize = True
                ALabel.Location = New System.Drawing.Point(iX, iY)
                ALabel.Name = col.Name + iCount.ToString()
                ALabel.Size = New System.Drawing.Size(53, 13)
                ALabel.TabIndex = 2 + iCount
                ALabel.Text = col.HeaderText
                Me.tpUsers.Controls.Add(ALabel)

                Dim c As DataGridViewCheckBoxColumn
                c = TryCast(col, DataGridViewCheckBoxColumn)
                If Not c Is Nothing Then
                    ACheckBox = New System.Windows.Forms.CheckBox
                    ACheckBox.DataBindings.Add(New System.Windows.Forms.Binding("CheckState", Me.V_usr_logBindingSource, col.Name, True))
                    ACheckBox.Location = New System.Drawing.Point(iX + 300, iY)
                    ACheckBox.Name = col.Name
                    ACheckBox.Size = New System.Drawing.Size(104, 24)
                    ACheckBox.TabIndex = 2 + iCount + 1
                    ACheckBox.Enabled = False
                    Me.tpUsers.Controls.Add(ACheckBox)
                Else
                    ATextBox = New System.Windows.Forms.TextBox
                    ATextBox.DataBindings.Add(New System.Windows.Forms.Binding("Text", Me.V_usr_logBindingSource, col.Name, True))
                    ATextBox.Location = New System.Drawing.Point(iX + 300, iY)
                    ATextBox.Name = col.Name
                    ATextBox.Size = New System.Drawing.Size(100, 20)
                    ATextBox.TabIndex = 2 + iCount + 1
                    ATextBox.ReadOnly = True
                    Me.tpUsers.Controls.Add(ATextBox)
                End If

                iY = iY + 30
                iCount = iCount + 2
            End If
        Next

        'add a label to identify the 'table' when adjusting columns headers etc.
        ALabel = New System.Windows.Forms.Label
        ALabel.AutoSize = True
        ALabel.Location = New System.Drawing.Point(iX, iY)
        ALabel.Name = "v_usr_log"
        ALabel.Size = New System.Drawing.Size(53, 13)
        ALabel.TabIndex = 2 + iCount
        ALabel.Text = "v_usr_log"
        'ALabel.ForeColor = Color.LightGray
        Me.tpUsers.Controls.Add(ALabel)

    End Sub

    Protected Overrides Sub frmLoad(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.frmLoad(sender, e)
        'vtble_column.dgtble.Visible = False

        vform_tble.AdjustPosition(vform)
        vform_tble_column__visibility.AdjustPosition(vform_tble)

        vform_grp2.AdjustPosition(vform)
        vform_grp2.SetdgLabel("The groups in which the selected form is used.")
        ' SetdgLabel2("The groups in which the selected form is used.")

        m_form_grp_log.AdjustPosition(vform_grp2)
        m_form_grp_log.SetdgLabel("The modifications made to the selected group and form.")
        vtble_column.AdjustPosition(vtble)
        vtble_column_header.AdjustPosition(vtble_column)


        If blnAll = True Then

            m_form_grp_log_all.SetdgLabel("All the modifications made to the groups.")
            m_usr_change_log_all.SetdgLabel("All the modifications made to the users.")

            vusr_log.AdjustPosition(vusr)
            m_usr_change_log.AdjustPosition(vusr)
            m_usr_change_log.SetdgLabel("The modifications made to the selected user.")

            vtxt_header.AdjustPosition(vtxt)
            vform_grp.AdjustPosition(vgrp)
            m_form_grp_groupbox.AdjustPosition(vform_grp)
            m_form_grp_log2.AdjustPosition(vform_grp)
            m_form_grp_log2.SetdgLabel("The modifications made to the selected group.")

            Dim p1 As Point = Label1.Location
            p1.X = M_form_grpDataGridView.Location.X + M_form_grpDataGridView.Width + 10
            Label1.Location = p1

            Dim p As Point = tbDescr.Location
            p.X = v_txt_headerDataGridView.Location.X + v_txt_headerDataGridView.Width + 10
            tbDescr.Location = p
            p = tbHeader.Location
            p.X = v_txt_headerDataGridView.Location.X
            tbHeader.Location = p
            tbHeader.Width = v_txt_headerDataGridView.Width + 10 + tbDescr.Width
        End If

        blnAllowUpdate = True
        FillTableAdapter()
        ' vtble_column.dgtble.Visible = False

        If blnAll = True Then
            Me.SetBindingNavigatorSource(M_langBindingSource)
            Dim size As Size = dgChild1.Size
            size.Width = tpUsers.Size.Width - dgChild1.Location.X - 10
            dgChild1.Size = size
        End If

        If iInitialFormHeight <> 0 Then
            TabControl1.Height = Me.Height - 120
            AdjustHeightsDataGridViews()
        End If
        If blnAll = True Then CreateuserStatisticsBoxes(M_usrDataGridView.Width)
        dgChild1.Visible = False


    End Sub

    'Prevent FillTableAdapter being called until all the grids have been defined.
    'Otherwise the RowEnter event fails because it is called before the columns have been defined.
    Public Overrides Sub FormIsActivated()
        'FillTableAdapter()
    End Sub

    Protected Overrides Sub FillTableAdapter()
        If blnAllowUpdate = True Then
            vtble.StoreRowIndexWithFocus()
            vtble_column.StoreRowIndexWithFocus()
            vtble_column_header.StoreRowIndexWithFocus()
            If blnAll = True Then
                Me.M_tbleTableAdapter.Fill(Me.TheDataSet.m_tble)
            Else
                Me.M_tbleTableAdapter.FillByForm(Me.TheDataSet.m_tble, strForm)
            End If
            vtble.ResetFocusRow()
            vtble_column.ResetFocusRow()
            vtble_column_header.ResetFocusRow()

            vform.StoreRowIndexWithFocus()
            vform_tble.StoreRowIndexWithFocus()
            vform_tble_column__visibility.StoreRowIndexWithFocus()

            If blnAll = True Then
                Me.M_formTableAdapter.FillByMenuIsForm(Me.TheDataSet.m_form)

                If tpUserChangeLog.Name = TabControl1.SelectedTab.Name Then
                    Me.M_usr_change_logTableAdapter.Fill(Me.TheDataSet.m_usr_change_log)
                Else
                    If Me.tpGrpFormLog.Name = TabControl1.SelectedTab.Name Then
                        Me.M_form_grp_logTableAdapter.Fill(Me.TheDataSet.m_form_grp_log)
                    End If
                End If

            Else
                If strForm.Length > 0 Then
                    Me.M_formTableAdapter.FillByFormMenuIsForm(Me.TheDataSet.m_form, strForm)
                End If
            End If

            vform.ResetFocusRow()
            vform_tble.ResetFocusRow()
            vform_tble_column__visibility.ResetFocusRow()

            If blnAll = True Then

                '20141230 put checks on the page being shown.
                If Me.tpUsers.Name = TabControl1.SelectedTab.Name Then
                    Me.vusr.Fill(Me.TheDataSet.m_usr)
                End If

                If Me.tpGroup.Name = TabControl1.SelectedTab.Name Then
                    vgrp.StoreRowIndexWithFocus()
                    vform_grp.StoreRowIndexWithFocus()
                    Me.M_grpTableAdapter.Fill(Me.TheDataSet.m_grp)
                    vgrp.ResetFocusRow()
                    vform_grp.ResetFocusRow()
                End If

                If Me.tpLanguage.Name = TabControl1.SelectedTab.Name Then
                    Me.vLang.Fill(Me.TheDataSet.m_lang)
                End If

                If Me.tpTexts.Name = TabControl1.SelectedTab.Name Then
                    vtxt.StoreRowIndexWithFocus()
                    vtxt_header.StoreRowIndexWithFocus()
                    M_txtTableAdapter.Fill(Me.TheDataSet.m_txt)
                    vtxt.ResetFocusRow()
                    vtxt_header.ResetFocusRow()
                End If

                If Me.tpFormat.Name = TabControl1.SelectedTab.Name Then
                    Me.vformat.Fill(Me.TheDataSet.m_format)
                End If

            End If

        End If

    End Sub

    Public Sub Fills()
        FillTableAdapter()
    End Sub

    Private Sub M_formDataGridView_RowEnter(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles M_formDataGridView.RowEnter
        If Me.M_form_tbleTableAdapter.FillBy(Me.TheDataSet.m_form_tble, M_formDataGridView.Rows(e.RowIndex).Cells("form").Value) = 0 Then
            Me.TheDataSet.m_form_tble_column__visibility.Clear()
        End If
        '20141229
        If Me.M_form_grpTableAdapter.FillByForm(Me.TheDataSet.m_form_grp, M_formDataGridView.Rows(e.RowIndex).Cells("form").Value) = 0 Then
            Me.TheDataSet.m_form_grp_log.Clear()
        End If
    End Sub

    '20141229
    Private Sub dgGrp_RowEnter(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgGrp.RowEnter

        Try
            'check if forms has the focus otherwise it will fire when groups has focus.
            If tpForms.Name = TabControl1.SelectedTab.Name Then
                Me.M_form_grp_logTableAdapter.FillByGrpForm(Me.TheDataSet.m_form_grp_log, Me.dgGrp.Rows(e.RowIndex).Cells("grp").Value, Me.dgGrp.Rows(e.RowIndex).Cells("form").Value)
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub M_form_tbleDataGridView_RowEnter(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles M_form_tbleDataGridView.RowEnter
        Me.M_form_tble_column__visibilityTableAdapter.FillBy(Me.TheDataSet.m_form_tble_column__visibility, _
                                                             M_form_tbleDataGridView.Rows(e.RowIndex).Cells("form").Value, _
                                                             M_form_tbleDataGridView.Rows(e.RowIndex).Cells("tble").Value)
    End Sub

    Private Sub M_grpDataGridView_RowEnter(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles M_grpDataGridView.RowEnter
        If Not M_grpDataGridView.Rows(e.RowIndex).Cells("grp").Value Is Nothing And tpGroup.Name = TabControl1.SelectedTab.Name Then
            Me.M_form_grpTableAdapter.FillByGrp(Me.TheDataSet.m_form_grp, M_grpDataGridView.Rows(e.RowIndex).Cells("grp").Value)

            '20141229
            Me.M_form_grp_logTableAdapter.FillByGrp(Me.TheDataSet.m_form_grp_log, Me.M_grpDataGridView.Rows(e.RowIndex).Cells("grp").Value)

        End If

    End Sub

    Private Sub M_form_grpDataGridView_RowEnter(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles M_form_grpDataGridView.RowEnter
        Try
            Me.M_form_grp_groupboxTableAdapter.FillBy(Me.TheDataSet.m_form_grp_groupbox, _
                                                      M_form_grpDataGridView.Rows(e.RowIndex).Cells("grp").Value, _
                                                      M_form_grpDataGridView.Rows(e.RowIndex).Cells("form").Value)

        Catch ex As Exception
            '   MsgBox(ex.Message)
        End Try

    End Sub


    Private Sub M_tbleDataGridView_RowEnter(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles M_tbleDataGridView.RowEnter
        Try
            If blnAllowUpdate = True Then
                Me.M_tble_columnTableAdapter.FillBy(Me.TheDataSet.m_tble_column, M_tbleDataGridView.Rows(e.RowIndex).Cells("tble").Value)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub M_tble_columnDataGridView_RowEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles M_tble_columnDataGridView.RowEnter
        Try
            If blnAllowUpdate = True Then
                ' MsgBox("71. vtble_column.dgtble.Visible = " + vtble_column.dgtble.Visible.ToString())
                Dim i As Integer = Me.V_tble_column_headerTableAdapter.FillByTbleColmn(Me.TheDataSet.v_tble_column_header, _
                M_tble_columnDataGridView.Rows(e.RowIndex).Cells("tble").Value, _
                M_tble_columnDataGridView.Rows(e.RowIndex).Cells("colmn").Value _
                )
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub m_txtDataGridView_RowEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles M_txtDataGridView.RowEnter
        Try
            V_txt_headerTableAdapter.FillByTxt(Me.TheDataSet.v_txt_header, _
            M_txtDataGridView.Rows(e.RowIndex).Cells("txt").Value)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub M_usrDataGridView_RowEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles M_usrDataGridView.RowEnter
        Dim cur As Cursor = Me.Cursor
        If Not cur = Cursors.WaitCursor Then
            Me.Cursor = Cursors.WaitCursor
        End If
        Try
            Me.V_usr_logTableAdapter.FillBy(Me.TheDataSet.v_usr_log, Application.ProductName, M_usrDataGridView.Rows(e.RowIndex).Cells("usr").Value)
            Me.M_usr_change_logTableAdapter.FillByUsr(Me.TheDataSet.m_usr_change_log, M_usrDataGridView.Rows(e.RowIndex).Cells("usr").Value)
        Catch ex As Exception
        End Try
        If Not cur = Cursors.WaitCursor Then
            Me.Cursor = cur
        End If

    End Sub

    Private Sub frmManage_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.FormClosing  '.Leave
        'Handles MyBase.Leave

        're-load the static tables because they may have been altered.
        statics.LoadData()

        '20120712 Call this to adjust the visibility if it was changed.
        MainDefs.MainForm.CheckFormVisibility()
    End Sub

    'This tabbed pages needs to override to go to tabpage control collection to find the dgvEnter.
    Protected Overrides Sub UpdateData()
        Dim dgv As dgvEnter
        Dim tp As TabPage
        Dim tc As TabControl
        'Look for dgvEnter controls.
        For Each c As Control In Controls
            tc = TryCast(c, TabControl)
            If Not tc Is Nothing Then
                For Each c2 As Control In tc.Controls
                    tp = TryCast(c2, TabPage)
                    If Not tp Is Nothing Then
                        For Each c3 As Control In tp.Controls
                            dgv = TryCast(c3, dgvEnter)
                            If Not dgv Is Nothing Then
                                Try
                                    dgv.UpdateData()
                                Catch ex As Exception
                                End Try
                            End If
                        Next
                    End If
                Next
            End If
        Next

    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Dim dgv As dgvEnter
        Dim tp As TabPage
        Dim tc As TabControl

        'Look for dgvEnter controls.
        tc = TryCast(sender, TabControl)
        tp = tc.SelectedTab
        If Not tp Is Nothing Then
            If tp.Name = "tpUsers" Then
                tsbUnBlock.Visible = True
            Else
                tsbUnBlock.Visible = False
            End If
            For Each c3 As Control In tp.Controls
                dgv = TryCast(c3, dgvEnter)
                If Not dgv Is Nothing Then
                    Try
                        Me.SetBindingNavigatorSource(dgv.DataSource)
                    Catch ex As Exception
                    End Try
                End If
            Next
        End If
        FillTableAdapter()
    End Sub

#End Region
#Region "UnBlock"
    Protected Sub tsbUnBlock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbUnBlock.Click
        Try
            If Not M_usrDataGridView.CurrentRow Is Nothing Then
                Dim strUser As String = M_usrDataGridView.CurrentRow.Cells("usr").Value
                If Not strUser Is Nothing Then
                    'Write a login message for the selected user setting logout flag to True so that the user
                    'is not shown as logged in. Only problem is if the user unblocks themselves because then they
                    'are shown as not logged in.
                    statics.UpdateUsrLog(strUser, True, False)
                    Me.V_usr_logTableAdapter.FillBy(Me.TheDataSet.v_usr_log, Application.ProductName, strUser)
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub
#End Region
#Region "Print"
    Protected Overrides Sub tsbPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim tp As TabPage = Me.TabControl1.SelectedTab
        Dim dgv As dgvEnter
        For Each c3 As Control In tp.Controls
            dgv = TryCast(c3, dgvEnter)
            If Not dgv Is Nothing Then
                Try
                    If dgv.Name = vLang.dg.Name Or _
                    dgv.Name = vusr.dg.Name Or _
                    dgv.Name = vform.dg.Name Or _
                    dgv.Name = vtble.dg.Name Or _
                    dgv.Name = vformat.dg.Name Or _
                    dgv.Name = vusr.dg.Name Or _
                    dgv.Name = vtxt.dg.Name _
                    Then
                        PrintToExcel(tp.Text, dgv)
                    Else
                        If dgv.Name = vgrp.dg.Name Then
                            PrintGroups()
                        End If
                    End If
                Catch ex As Exception
                End Try
            End If
        Next
    End Sub

    'Print the groups with the forms.
    Protected Sub PrintGroups()
        Dim cur = Me.Cursor
        Try
            Me.Cursor = Cursors.WaitCursor
            Dim pr As New ExcelInterface.XMLExcelInterface("")  '"" is the Network path so we go local.
            pr.OpenExcelBook(ExcelInterface.Paths.Local, "", "Groups", False, My.Settings.XMLTemplate)
            pr.NewSheet("Groups", _
                "&amp;LUsing data from " + "Groups" + "." + "&amp;CPrinted on &amp;D &amp;T. " + "&amp;RPage &amp;P of &amp;N", _
                True, "Groups")
            pr.WriteColumnWidths(M_form_grpDataGridView, MainDefs.DONOTPRINT, False, 0)

            Dim sD As System.Data.DataRowView
            Dim aRow As TheDataSet.m_grpRow
            Me.M_grpBindingSource.MoveFirst()
            Dim iPosition As Integer = -1
            Dim iFirstPosition As Integer
            iFirstPosition = M_grpBindingSource.Position


            Do While iPosition <> M_grpBindingSource.Position
                iPosition = M_grpBindingSource.Position
                sD = Me.M_grpBindingSource.Current
                aRow = sD.Row

                pr.WriteStringToExcel("Group#" + aRow.grp + "#", ExcelInterface.ExcelStringFormats.Bold10)
                pr.WriteDataGrid(M_form_grpDataGridView, MainDefs.DONOTPRINT, False, 0, False, False, True)
                pr.WriteStringToExcel("#", ExcelInterface.ExcelStringFormats.Bold10)
                Me.M_grpBindingSource.MoveNext()

            Loop
            M_grpBindingSource.Position = iFirstPosition
            pr.CloseExcelBook()
            NAR(pr) ' = Nothing
        Catch ex As Exception
        End Try
        Me.Cursor = cur
    End Sub
#End Region
    'Start dialog for selecting forms for groups.
    Private Sub M_grpDataGridView_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles M_grpDataGridView.KeyPress
        Try
            Dim strColumnName As String = M_grpDataGridView.Columns(M_grpDataGridView.CurrentCell.ColumnIndex).Name
            If strColumnName = "grp" Then

                '20200219 Do not allow dialog DialogSelectForms to be shown if the Security form is Read only.
                If blnRO = False Then
                    If M_grpDataGridView.CurrentRow.Cells(strColumnName).ToString.Length > 0 Then

                        Dim D As New DialogSelectForms(MainDefs, M_grpDataGridView.CurrentCell.Value, Me.M_form_grpBindingSource) ', Me.M_form_grpTableAdapter.GetDataByGrp(M_grpDataGridView.CurrentCell.Value))
                        If D.ShowDialog() = DialogResult.OK Then

                            '20131031 This was loading all the forms and it should be by group.
                            Me.M_form_grpTableAdapter.FillByGrp(Me.TheDataSet.m_form_grp, M_grpDataGridView.CurrentRow.Cells("grp").Value)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Protected Overrides Sub frm_Layout(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles MyBase.Layout
        If TestActiveMDIChild() = True Then

            TabControl1.Height = Me.Height - 120
            Dim l As Point
            l = Me.TabControl1.Location
            l.Y = 30
            Me.TabControl1.Location = l
            AdjustHeightsDataGridViews()
        End If
    End Sub
End Class
