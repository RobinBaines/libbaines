<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SQLBrowser
    Inherits Utilities.frmStandard

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.TVUses = New System.Windows.Forms.TreeView()
        Me.TVUsedBy = New System.Windows.Forms.TreeView()
        Me.tbSearch = New System.Windows.Forms.TextBox()
        Me.tbSchema = New System.Windows.Forms.TextBox()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.gbUses = New System.Windows.Forms.GroupBox()
        Me.gbUsedBy = New System.Windows.Forms.GroupBox()
        Me.EntityCommand1 = New System.Data.Entity.Core.EntityClient.EntityCommand()
        CType(Me.BindingNavigator, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        Me.gbUses.SuspendLayout()
        Me.gbUsedBy.SuspendLayout()
        '
        'BindingNavigator
        '
        Me.BindingNavigator.Size = New System.Drawing.Size(1728, 25)
        '
        'tbSearch
        '
        Me.tbSearch.Location = New System.Drawing.Point(185, 32)
        Me.tbSearch.Name = "tbSearch"
        Me.tbSearch.Size = New System.Drawing.Size(153, 23)
        Me.tbSearch.TabIndex = 3

        '
        'tbSchema
        '
        Me.tbSchema.Location = New System.Drawing.Point(30, 32)
        Me.tbSchema.Name = "tbSchema"
        Me.tbSchema.Size = New System.Drawing.Size(153, 20)
        Me.tbSchema.TabIndex = 2

        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(340, 32)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(75, 23)
        Me.btnSearch.TabIndex = 4
        Me.btnSearch.Text = "Search"
        Me.btnSearch.UseMnemonic = False
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'TreeView1
        '
        Me.TVUses.Location = New System.Drawing.Point(10, 20)
        Me.TVUses.Name = "TVUses"
        Me.TVUses.Size = New System.Drawing.Size(467, 370)
        Me.TVUses.TabIndex = 0

        '
        'TVUsedBy
        '
        Me.TVUsedBy.Location = New System.Drawing.Point(10, 20)
        Me.TVUsedBy.Name = "TVUsedBy"
        Me.TVUsedBy.Size = New System.Drawing.Size(467, 370)
        Me.TVUsedBy.TabIndex = 10

        '
        'gbUses
        '
        'Me.gbUses.Controls.Add(Me.tbSearch)
        'Me.gbUses.Controls.Add(Me.btnSearch)
        Me.gbUses.Controls.Add(Me.TVUses)
        Me.gbUses.Location = New System.Drawing.Point(12, 60)
        Me.gbUses.Name = "gbUses"
        Me.gbUses.Size = New System.Drawing.Size(483, 296)
        Me.gbUses.TabIndex = 5
        Me.gbUses.TabStop = False
        Me.gbUses.Text = "Uses"

        '
        'gbUsedBy
        '
        'Me.gbUsedBy.Controls.Add(Me.tbSearch)
        'Me.gbUsedBy.Controls.Add(Me.btnSearch)
        Me.gbUsedBy.Controls.Add(Me.TVUsedBy)
        Me.gbUsedBy.Location = New System.Drawing.Point(12, 330)
        Me.gbUsedBy.Name = "gbUsedBy"
        Me.gbUsedBy.Size = New System.Drawing.Size(483, 296)
        Me.gbUsedBy.TabIndex = 6
        Me.gbUsedBy.TabStop = False
        Me.gbUsedBy.Text = "UsedBy"

        '
        'EntityCommand1
        '
        Me.EntityCommand1.CommandTimeout = 0
        Me.EntityCommand1.CommandTree = Nothing
        Me.EntityCommand1.Connection = Nothing
        Me.EntityCommand1.EnablePlanCaching = True
        Me.EntityCommand1.Transaction = Nothing
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = False
        Me.ClientSize = New System.Drawing.Size(1728, 828)
        Me.Controls.Add(Me.gbUses)
        Me.Controls.Add(Me.gbUsedBy)

        Me.Controls.Add(tbSearch)
        Me.Controls.Add(tbSchema)
        Me.Controls.Add(btnSearch)

        Me.Location = New System.Drawing.Point(0, 0)
        Me.Name = "Form2"
        Me.Text = "Form2"

        'Me.Controls.SetChildIndex(Me.TVUses, 0)
        'Me.Controls.SetChildIndex(Me.TVUsedBy, 0)
        Me.Controls.SetChildIndex(Me.BindingNavigator, 0)

        CType(Me.BindingNavigator, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbUses.ResumeLayout(False)
        Me.gbUses.PerformLayout()
        Me.gbUsedBy.ResumeLayout(False)
        Me.gbUsedBy.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub


    Friend WithEvents TVUses As TreeView
    Friend WithEvents TVUsedBy As TreeView
    Friend WithEvents SqlCommand1 As Microsoft.Data.SqlClient.SqlCommand
    Friend WithEvents EntityCommand1 As Entity.Core.EntityClient.EntityCommand
    Friend WithEvents tbSearch As TextBox
    Friend WithEvents tbSchema As TextBox

    Friend WithEvents btnSearch As Button
    Friend WithEvents gbUses As GroupBox
    Friend WithEvents gbUsedBy As GroupBox
End Class
