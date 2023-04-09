<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProperties
    Inherits frmStandard

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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnSave = New System.Windows.Forms.Button
        Me.cbQuality = New System.Windows.Forms.CheckBox
        Me.lLiveDataSource = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lTestDataSource = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.lConnectionString = New System.Windows.Forms.Label
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog
        Me.Label4 = New System.Windows.Forms.Label
        Me.lLiveCatalog = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.lTestCatalog = New System.Windows.Forms.Label
        Me.cbEnableAudio = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(143, 276)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(56, 28)
        Me.btnSave.TabIndex = 6
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'cbQuality
        '
        Me.cbQuality.AutoSize = True
        Me.cbQuality.Location = New System.Drawing.Point(143, 167)
        Me.cbQuality.Margin = New System.Windows.Forms.Padding(2)
        Me.cbQuality.Name = "cbQuality"
        Me.cbQuality.Size = New System.Drawing.Size(96, 17)
        Me.cbQuality.TabIndex = 7
        Me.cbQuality.Text = "Test Database"
        Me.cbQuality.UseVisualStyleBackColor = True
        '
        'lLiveDataSource
        '
        Me.lLiveDataSource.AutoSize = True
        Me.lLiveDataSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lLiveDataSource.Location = New System.Drawing.Point(140, 51)
        Me.lLiveDataSource.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lLiveDataSource.Name = "lLiveDataSource"
        Me.lLiveDataSource.Size = New System.Drawing.Size(106, 13)
        Me.lLiveDataSource.TabIndex = 10
        Me.lLiveDataSource.Text = "Live Data Source"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 51)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 13)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Live Data Source"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 79)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(91, 13)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Test Data Source"
        '
        'lTestDataSource
        '
        Me.lTestDataSource.AutoSize = True
        Me.lTestDataSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lTestDataSource.Location = New System.Drawing.Point(140, 79)
        Me.lTestDataSource.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lTestDataSource.Name = "lTestDataSource"
        Me.lTestDataSource.Size = New System.Drawing.Size(107, 13)
        Me.lTestDataSource.TabIndex = 15
        Me.lTestDataSource.Text = "Test Data Source"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 123)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(91, 13)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Connection String"
        '
        'lConnectionString
        '
        Me.lConnectionString.AutoSize = True
        Me.lConnectionString.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lConnectionString.Location = New System.Drawing.Point(141, 123)
        Me.lConnectionString.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lConnectionString.Name = "lConnectionString"
        Me.lConnectionString.Size = New System.Drawing.Size(108, 13)
        Me.lConnectionString.TabIndex = 16
        Me.lConnectionString.Text = "Connection String"
        '
        'ColorDialog1
        '
        Me.ColorDialog1.AnyColor = True
        Me.ColorDialog1.FullOpen = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(444, 51)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(66, 13)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Live Catalog"
        '
        'lLiveCatalog
        '
        Me.lLiveCatalog.AutoSize = True
        Me.lLiveCatalog.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lLiveCatalog.Location = New System.Drawing.Point(541, 51)
        Me.lLiveCatalog.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lLiveCatalog.Name = "lLiveCatalog"
        Me.lLiveCatalog.Size = New System.Drawing.Size(78, 13)
        Me.lLiveCatalog.TabIndex = 18
        Me.lLiveCatalog.Text = "Live Catalog"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(444, 79)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(67, 13)
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "Test Catalog"
        '
        'lTestCatalog
        '
        Me.lTestCatalog.AutoSize = True
        Me.lTestCatalog.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lTestCatalog.Location = New System.Drawing.Point(541, 79)
        Me.lTestCatalog.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lTestCatalog.Name = "lTestCatalog"
        Me.lTestCatalog.Size = New System.Drawing.Size(79, 13)
        Me.lTestCatalog.TabIndex = 20
        Me.lTestCatalog.Text = "Test Catalog"
        '
        'cbEnableAudio
        '
        Me.cbEnableAudio.AutoSize = True
        Me.cbEnableAudio.Location = New System.Drawing.Point(143, 201)
        Me.cbEnableAudio.Name = "cbEnableAudio"
        Me.cbEnableAudio.Size = New System.Drawing.Size(89, 17)
        Me.cbEnableAudio.TabIndex = 22
        Me.cbEnableAudio.Text = "Enable Audio"
        Me.cbEnableAudio.UseVisualStyleBackColor = True
        '
        'frmProperties
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(932, 466)
        Me.Controls.Add(Me.cbEnableAudio)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.lTestCatalog)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.lLiveCatalog)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.lConnectionString)
        Me.Controls.Add(Me.lTestDataSource)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lLiveDataSource)
        Me.Controls.Add(Me.cbQuality)
        Me.Controls.Add(Me.btnSave)
        Me.Name = "frmProperties"
        Me.Text = "Properties"
        Me.Controls.SetChildIndex(Me.btnSave, 0)
        Me.Controls.SetChildIndex(Me.cbQuality, 0)
        Me.Controls.SetChildIndex(Me.lLiveDataSource, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.lTestDataSource, 0)
        Me.Controls.SetChildIndex(Me.lConnectionString, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.lLiveCatalog, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.lTestCatalog, 0)
        Me.Controls.SetChildIndex(Me.Label5, 0)
        Me.Controls.SetChildIndex(Me.cbEnableAudio, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents cbQuality As System.Windows.Forms.CheckBox
    Friend WithEvents lLiveDataSource As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lTestDataSource As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lConnectionString As System.Windows.Forms.Label
    Friend WithEvents ColorDialog1 As System.Windows.Forms.ColorDialog
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lLiveCatalog As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lTestCatalog As System.Windows.Forms.Label
    Friend WithEvents cbEnableAudio As System.Windows.Forms.CheckBox
End Class
