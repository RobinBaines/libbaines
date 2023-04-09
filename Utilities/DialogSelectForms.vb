'------------------------------------------------
'Name: Module DialogSelectForms.vb
'Function: Dialog to select forms from frmManage.
'Copyright Robin Baines 2010. All rights reserved.
'20141229 Modified by using a temporary table, m_form_grp_temp, to update m_form_grp so that updating fills the log in the expected way.
'20151107 moved try catch because p_update_m_form_grp should not be called if p_insert_m_form_grp_temp fails.
'and because it might have failed the last time empty the m_form_grp_temp table.
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms

Public Class DialogSelectForms
    Friend WithEvents M_formTableAdapter As TheDataSetTableAdapters.m_formTableAdapter
    Friend WithEvents M_form_grpTableAdapter As TheDataSetTableAdapters.m_form_grpTableAdapter
    Dim vform_grp As TheDataSet_m_form_grp
    Dim strConnection As String

    Private Function IsInGrpForm(ByVal bsDestination As System.Windows.Forms.BindingSource, ByVal strForm As String) As Boolean
        Dim blnRet As Boolean = False
        Dim sD As System.Data.DataRowView
        Dim aRowD As TheDataSet.m_form_grpRow
        For Each sD In bsDestination.List()
            aRowD = sD.Row
            If aRowD.form = strForm Then
                blnRet = True
                Exit For
            End If
        Next
        Return blnRet
    End Function

    Public Sub New(ByVal MainDef As MainDefinitions, ByVal strGrp As String, ByVal bsDestination As System.Windows.Forms.BindingSource)  ', ByVal form_grp As TheDataSet.m_form_grpDataTable)
        InitializeComponent()

        Me.M_formTableAdapter = New TheDataSetTableAdapters.m_formTableAdapter
        M_formTableAdapter.Connection.ConnectionString = MainDef.GetConnectionString()
        strConnection = MainDef.GetConnectionString()
        Dim formTable As TheDataSet.m_formDataTable = Me.M_formTableAdapter.GetDataByTopLevelMenus()

        tbGroup.Text = strGrp
        tbGroup.ReadOnly = True

        dgM_form_grp.AllowUserToAddRows = False
        dgM_form_grp.AllowUserToDeleteRows = False

        Dim aRowD2 As TheDataSet.m_form_grpRow
        Dim sD2 As System.Data.DataRowView
        For Each sD2 In bsDestination.List()
            aRowD2 = sD2.Row
            Dim iRow As Integer = dgM_form_grp.Rows.Add()
            dgM_form_grp.Rows(iRow).Cells("form").Value = aRowD2.form
            dgM_form_grp.Rows(iRow).Cells("ro").Value = aRowD2.RO
        Next

        'Source
        ListBox1.Sorted = True
        ListBox1.BeginUpdate()
        Dim aRowS As TheDataSet.m_formRow
        For Each aRowS In formTable
            If IsInGrpForm(bsDestination, aRowS.form) Then
            Else
                ListBox1.Items.Add(aRowS.form)
            End If
        Next
        If ListBox1.Items.Count > 0 Then
            ListBox1.SelectedIndex = 0
        End If
        ListBox1.EndUpdate()
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.Cursor = Cursors.WaitCursor
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        M_form_grpTableAdapter = New TheDataSetTableAdapters.m_form_grpTableAdapter
        M_form_grpTableAdapter.Connection.ConnectionString = strConnection ' M_formTableAdapter.Connection.ConnectionString()



        '20151107 moved try catch because p_update_m_form_grp should not be called if p_insert_m_form_grp_temp fails.
        'and because it might have failed the last time empty the m_form_grp_temp table.
        M_form_grpTableAdapter.Delete_m_form_grp_temp()
        Try
            'insert all the entries into a temporary table called m_form_grp_temp
            For Each r As DataGridViewRow In dgM_form_grp.Rows
                Dim blnRO As Boolean = r.Cells("ro").Value
                M_form_grpTableAdapter.p_insert_m_form_grp_temp(tbGroup.Text, r.Cells("form").Value, blnRO)
           
            Next

        'then use the temporary table, m_form_grp_temp, to update m_form_grp. 
            M_form_grpTableAdapter.p_update_m_form_grp(tbGroup.Text)

        Catch ex As Exception
        End Try
        Me.Cursor = Cursors.Default
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Try
            If ListBox1.Items.Count > 0 Then
                If Not ListBox1.SelectedItem Is Nothing Then
                    Dim iRow As Integer = dgM_form_grp.Rows.Add()
                    dgM_form_grp.Rows(iRow).Cells("form").Value = ListBox1.SelectedItem
                    dgM_form_grp.Rows(iRow).Cells("ro").Value = 0
                    ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
                    If Not ListBox1.Items.Count = 0 Then
                        ListBox1.SelectedIndex = 0
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnAddAll2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAll.Click
        Try
            Dim iRow As Integer
            If ListBox1.Items.Count > 0 Then
                For Each item As String In ListBox1.Items
                    iRow = dgM_form_grp.Rows.Add()
                    dgM_form_grp.Rows(iRow).Cells("form").Value = item  'ListBox1.SelectedItem
                    dgM_form_grp.Rows(iRow).Cells("ro").Value = 0
                Next
                ListBox1.Items.Clear()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        Try
            If dgM_form_grp.Rows.Count > 0 Then
                ListBox1.Items.Add(dgM_form_grp.CurrentRow.Cells("form").Value())
                dgM_form_grp.Rows.RemoveAt(dgM_form_grp.CurrentRow.Index)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnRemoveAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveAll.Click
        Try
            If dgM_form_grp.Rows.Count > 0 Then
                For Each r As DataGridViewRow In dgM_form_grp.Rows
                    ListBox1.Items.Add(r.Cells("form").Value)
                Next
                dgM_form_grp.Rows.Clear()
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class
