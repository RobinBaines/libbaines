'------------------------------------------------
'Name: Module frmUsrLog2.vb
'Function: Show the user log.
'Copyright Robin Baines 2010. All rights reserved.
'Notes: The following steps describe how to create a new form based on frmStandard and how to define a datagridview
'based on dgvEnter and dgColumns.
'The resulting form and datagridview use the utilities functionality with column design, segregation of duty, excel export, 
'timer calls from Main and events from the database (see frmPollingExample.vb).

' The following steps are a guideline and alternative approaches are possible.
'1. create a new form using the solution explorer. Give it a name: frmUsrLog2 in this example.
'2. open frmUsrLog2.Designer.vb and change Inherits statement to 
' Inherits frmStandard
'3. open the toolbox and add a dgvEnter datagridview to the form. Optional is to give the dgvEnter a name for example dgParent.
'4. open frmUsrLog2.vb
'5. open for example, Utilitites.frmAppLog, as an example of code behind the form and paste into frmUsrLog2.vb
'6. if the form uses new views or tables to show data add these to a new or existing Dataset. Save the modified Dataset.
'7. create the dgvEnter classes for new views using the GENUI (github) project from the TestApp Dataset file. 

'-------------The GenUI md file--------------------------------
'The GenUI executable Is used To generate classes for views defined in the datasets in a visual studio project.
'For example

'run\ genui BAINESLENOVO TestDb c:\Projects\AppsLib\libbaines\TestApp\TestApp\ c:\projects\GenUI\OUTPUT\
'run\ genui [SQL Server] [database] [dataset files folder] [output folder]

'Looks for tables and views in Dataset xsd files in a project folder: c : \Projects\AppsLib\libbaines\TestApp\TestApp\
'Looks up the definition of the views on SQL Server BAINESLENOVO in database TestDb and generates code using the SQL Server api.
'This only works correctly if a single table or view is used in a tableadapter in the dataset. 
'(Any JOINs are defined in SQL Server views.)

'For example the above genui generates code gen_TheDataSet_v_usr_log.vb

'Public Class TheDataSet_v_usr_log
'    Inherits Utilities.dgColumns
'    'columns of the datagrid.
'    'textboxes for the filter textboxes above each column.
'---------------------------------------------

'In this example an existing view is used and the dgvEnter class TheDataSet_v_usr_log is already present.
'9. Open the form designer and the Data Sources pane. Drop the TheDataSet.v_usr_log on to the dgvEnter.
'This creates the V_usr_logTableAdapter and V_usr_logBindingSource.
'10. However it also creates the columns in the grid. Remove these columns by deleting dgvEnter property
'DataSource 'V_usr_logBindingSource' replacing it with None. (The calls in gen_TheDataSet_v_usr_log.vb generate the columns.)
'Also remove the Load sub routine which is added automatically at the bottom of the code infrmUsrLog2.vb in step 9.. 
'This will include the Fill call: Me.V_usr_logTableAdapter.Fill(Me.TheDataSet.v_usr_log) which is called from 
'Protected Overrides Sub FillTableAdapter(). 
'11. clean up the code, see below.
'12. open the form from a menu in the mainform, see TestAppMain.vb:
'Private Sub tsbfrmTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsbfrmTest.Click
'    If blnBringToFrontIfExists(Me, sender.ToolTipText) = False Then
'        ShowAForm(Me, New frmUsrLog2(sender, sender.ToolTipText, MainDefs), sender.Text, sender.ToolTipText)
'    End If
'End Sub
'Modifications: 
'------------------------------------------------
Imports Utilities

Public Class frmUsrLog2
#Region "New"
    Dim vParent As TheDataSet_v_usr_log
    Public Sub New()

        MyBase.New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

    End Sub
    Public Sub New(ByVal tsb As ToolStripItem _
               , ByVal strSecurityName As String, ByVal _MainDefs As MainDefinitions)

        MyBase.New(tsb, strSecurityName, _MainDefs)
        InitializeComponent()
        vParent = New TheDataSet_v_usr_log(strSecurityName, V_usr_logBindingSource, dgParent, V_usr_logTableAdapter,
            Me.TheDataSet,
            Me.components,
            MainDefs, True, Controls, Me, True)

        SetBindingNavigatorSource(V_usr_logBindingSource)
        Me.SwitchOffPrintDetail()
        Me.SwitchOffPrint()
        Me.SwitchOffUpdate()

    End Sub

#End Region

#Region "Load"

    Protected Overrides Sub frmLoad(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles MyBase.Load
        MyBase.frmLoad(sender, e)
        blnAllowUpdate = True
        FillTableAdapter()
    End Sub
    Protected Overrides Sub FillTableAdapter()
        MyBase.FillTableAdapter()
        vParent.StoreRowIndexWithFocus()
        Me.V_usr_logTableAdapter.Fill(Me.TheDataSet.v_usr_log)
        vParent.ResetFocusRow()

    End Sub
#End Region

#Region "Scroll_Resize"
    'The datagrids are re-sized when the form re-sizes. But this causes problems if the re-size fires when the window is not the ActiveMDIChild.
    'This occurs if the Ctrl-tab combination is used to cycle through the windows of the application followed by an Alt.
    'This was also a problem when the form was re-writing when semaphore fired.
    'Solution is only to re-size when the form is the ActiveMDIChild.
    'Tried also to check on the windowstate so Resize occurs if the windowstate is not maximised.
    'But it appears that the windowstate is Normal if is maximized but is not the ActiveMDIChild.
    'Private Sub frm_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
    Protected Overrides Sub frm_Layout(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LayoutEventArgs) 'Handles MyBase.Layout
        MyBase.frm_Layout(sender, e)
        If TestActiveMDIChild() = True Then
            If Not vParent Is Nothing Then
                vParent.SetHeight(Me.ClientRectangle.Height) ' dgParent.Height = Me.Height - 40 - dgParent.Location.Y
            End If
        End If
    End Sub

#End Region

End Class