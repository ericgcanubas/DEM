﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmSetup
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
        Me.rdBranch = New System.Windows.Forms.RadioButton()
        Me.rdMainOffice = New System.Windows.Forms.RadioButton()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.txtDatabase = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'rdBranch
        '
        Me.rdBranch.AutoSize = True
        Me.rdBranch.Location = New System.Drawing.Point(12, 12)
        Me.rdBranch.Name = "rdBranch"
        Me.rdBranch.Size = New System.Drawing.Size(59, 17)
        Me.rdBranch.TabIndex = 0
        Me.rdBranch.TabStop = True
        Me.rdBranch.Text = "Branch"
        Me.rdBranch.UseVisualStyleBackColor = True
        '
        'rdMainOffice
        '
        Me.rdMainOffice.AutoSize = True
        Me.rdMainOffice.Location = New System.Drawing.Point(103, 12)
        Me.rdMainOffice.Name = "rdMainOffice"
        Me.rdMainOffice.Size = New System.Drawing.Size(79, 17)
        Me.rdMainOffice.TabIndex = 1
        Me.rdMainOffice.TabStop = True
        Me.rdMainOffice.Text = "Main Office"
        Me.rdMainOffice.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(107, 156)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(12, 66)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(173, 20)
        Me.txtServer.TabIndex = 3
        '
        'txtDatabase
        '
        Me.txtDatabase.Location = New System.Drawing.Point(12, 120)
        Me.txtDatabase.Name = "txtDatabase"
        Me.txtDatabase.Size = New System.Drawing.Size(173, 20)
        Me.txtDatabase.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 47)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(44, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Server :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 104)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Database :"
        '
        'FrmSetup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(197, 189)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtDatabase)
        Me.Controls.Add(Me.txtServer)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.rdMainOffice)
        Me.Controls.Add(Me.rdBranch)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmSetup"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "First Setup"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents rdBranch As RadioButton
    Friend WithEvents rdMainOffice As RadioButton
    Friend WithEvents btnSave As Button
    Friend WithEvents txtServer As TextBox
    Friend WithEvents txtDatabase As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
End Class
