'------------------------------------------------
'Name: Module frmProperties.vb.
'Function: Form for switching between test and live databases.
'Copyright Robin Baines 2008. All rights reserved.
'Created 6/13/2008 12:00:00 AM.
'Notes: 
'Modifications:
'------------------------------------------------
Imports System.Windows.Forms
Imports Utilities

Public Class frmProperties
    Private Sub frmProperties_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If Me.MainDefs.MainForm.GetQualityDataBase() <> cbQuality.Checked Then
            If MsgBox("You have not saved alterations. Are you certain you want to close this form?", MsgBoxStyle.YesNo) <> MsgBoxResult.Yes Then
                e.Cancel = True
            End If
        End If
    End Sub
#Region "New"
    Public Sub New(ByVal tsb As ToolStripItem _
        , ByVal strSecurityName As String _
        , ByVal _MainDefs As MainDefinitions)

        MyBase.New(tsb, strSecurityName, _MainDefs)
        InitializeComponent()

        'disable the selective activation of child controls.
        'default is disabled but is correct to do it explicitly.
        Me.Enable_ChkLevel_On_ChildControls = False
        Init()
        Me.btnSave.Enabled = False
    End Sub
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        Init()

    End Sub
   
    Private Sub Init()
        Me.BindingNavigatorInvisible()
        cbQuality.Checked = Me.MainDefs.MainForm.GetQualityDataBase()
        Me.lLiveDataSource.Text = Me.MainDefs.MainForm.GetDataSourceLive()
        Me.lTestDataSource.Text = Me.MainDefs.MainForm.GetDataSourceQuality()
        Me.lLiveCatalog.Text = Me.MainDefs.MainForm.GetCatalogLive()
        Me.lTestCatalog.Text = Me.MainDefs.MainForm.GetCatalogQuality()
        Me.cbEnableAudio.Checked = Me.MainDefs.MainForm.GetEnableAudio()

        '20100916 Decided to disable the connection string for SFC.
        lConnectionString.Text = Me.MainDefs.MainForm.GetConnectionString(cbQuality.Checked)
        lConnectionString.Visible = False
        Label3.Visible = False

        'if no test database is defined then disable the check box.
        If lTestDataSource.Text.Length = 0 Then
            cbQuality.Enabled = False
            'Me.btnSave.Enabled = False
        End If
    End Sub
#End Region
#Region "Save"
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Me.MainDefs.MainForm.GetQualityDataBase() <> cbQuality.Checked Then
            Me.MainDefs.MainForm.SetQualityDataBase(cbQuality.Checked)
            Me.MainDefs.MainForm.CloseAllForms()
            Me.MainDefs.MainForm.Init()
        End If

        If Me.MainDefs.MainForm.GetEnableAudio() <> cbEnableAudio.Checked Then
            Me.MainDefs.MainForm.SetEnableAudio(cbEnableAudio.Checked)
        End If
    End Sub
#End Region
#Region "Validation"
    Private Sub cbQuality_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbQuality.CheckedChanged
        Dim strDataSource As String = Me.MainDefs.MainForm.GetDataSourceLive
        lConnectionString.Text = MainDefs.MainForm.GetConnectionString(cbQuality.Checked)
        Me.btnSave.Enabled = True
    End Sub
    Private Sub cbEnableAudio_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbEnableAudio.CheckedChanged
        Me.btnSave.Enabled = True
    End Sub
#End Region
End Class