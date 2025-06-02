<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMainSetup
    Inherits System.Windows.Forms.Form

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
        Me.LvLIst = New System.Windows.Forms.ListView()
        Me.Counter = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.BtnAdded = New System.Windows.Forms.Button()
        Me.txtCounter = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.chkGenerateType = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'LvLIst
        '
        Me.LvLIst.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Counter, Me.ColumnHeader1, Me.ColumnHeader2})
        Me.LvLIst.FullRowSelect = True
        Me.LvLIst.GridLines = True
        Me.LvLIst.HideSelection = False
        Me.LvLIst.Location = New System.Drawing.Point(10, 42)
        Me.LvLIst.Name = "LvLIst"
        Me.LvLIst.Size = New System.Drawing.Size(447, 273)
        Me.LvLIst.TabIndex = 6
        Me.LvLIst.UseCompatibleStateImageBehavior = False
        Me.LvLIst.View = System.Windows.Forms.View.Details
        '
        'Counter
        '
        Me.Counter.Text = "Counter"
        Me.Counter.Width = 108
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Last Upload"
        Me.ColumnHeader1.Width = 130
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Reference"
        Me.ColumnHeader2.Width = 160
        '
        'BtnAdded
        '
        Me.BtnAdded.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.BtnAdded.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnAdded.Location = New System.Drawing.Point(152, 13)
        Me.BtnAdded.Name = "BtnAdded"
        Me.BtnAdded.Size = New System.Drawing.Size(60, 23)
        Me.BtnAdded.TabIndex = 1
        Me.BtnAdded.Text = "Add"
        Me.BtnAdded.UseVisualStyleBackColor = False
        '
        'txtCounter
        '
        Me.txtCounter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCounter.Location = New System.Drawing.Point(63, 13)
        Me.txtCounter.MaxLength = 3
        Me.txtCounter.Name = "txtCounter"
        Me.txtCounter.Size = New System.Drawing.Size(84, 20)
        Me.txtCounter.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Counter :"
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDelete.Location = New System.Drawing.Point(218, 13)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(60, 23)
        Me.btnDelete.TabIndex = 4
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = False
        '
        'btnRefresh
        '
        Me.btnRefresh.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.btnRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnRefresh.Location = New System.Drawing.Point(397, 13)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(60, 23)
        Me.btnRefresh.TabIndex = 5
        Me.btnRefresh.Text = "Refresh"
        Me.btnRefresh.UseVisualStyleBackColor = False
        '
        'chkGenerateType
        '
        Me.chkGenerateType.AutoSize = True
        Me.chkGenerateType.Location = New System.Drawing.Point(12, 321)
        Me.chkGenerateType.Name = "chkGenerateType"
        Me.chkGenerateType.Size = New System.Drawing.Size(178, 17)
        Me.chkGenerateType.TabIndex = 7
        Me.chkGenerateType.Text = "Download always Optimize Data"
        Me.chkGenerateType.UseVisualStyleBackColor = True
        '
        'FrmMainSetup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(464, 370)
        Me.Controls.Add(Me.chkGenerateType)
        Me.Controls.Add(Me.LvLIst)
        Me.Controls.Add(Me.btnRefresh)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtCounter)
        Me.Controls.Add(Me.BtnAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmMainSetup"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Main Setup"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnAdded As Button
    Friend WithEvents txtCounter As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnDelete As Button
    Friend WithEvents btnRefresh As Button
    Friend WithEvents Counter As ColumnHeader
    Friend WithEvents ColumnHeader1 As ColumnHeader
    Friend WithEvents ColumnHeader2 As ColumnHeader
    Friend WithEvents LvLIst As ListView
    Friend WithEvents chkGenerateType As CheckBox
End Class
