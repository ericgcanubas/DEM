﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
        Me.lblMainLoading = New System.Windows.Forms.Label()
        Me.pbMainLoading = New System.Windows.Forms.ProgressBar()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblBranchLoading = New System.Windows.Forms.Label()
        Me.pbBranchLoading = New System.Windows.Forms.ProgressBar()
        Me.lblClose = New System.Windows.Forms.LinkLabel()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblMainLoading
        '
        Me.lblMainLoading.BackColor = System.Drawing.Color.Transparent
        Me.lblMainLoading.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblMainLoading.Location = New System.Drawing.Point(31, 235)
        Me.lblMainLoading.Name = "lblMainLoading"
        Me.lblMainLoading.Size = New System.Drawing.Size(287, 23)
        Me.lblMainLoading.TabIndex = 3
        Me.lblMainLoading.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pbMainLoading
        '
        Me.pbMainLoading.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.pbMainLoading.Location = New System.Drawing.Point(31, 209)
        Me.pbMainLoading.Name = "pbMainLoading"
        Me.pbMainLoading.Size = New System.Drawing.Size(287, 23)
        Me.pbMainLoading.Step = 1
        Me.pbMainLoading.TabIndex = 2
        '
        'btnImport
        '
        Me.btnImport.BackColor = System.Drawing.Color.Moccasin
        Me.btnImport.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnImport.Font = New System.Drawing.Font("Tahoma", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnImport.ForeColor = System.Drawing.Color.Red
        Me.btnImport.Location = New System.Drawing.Point(375, 167)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(288, 36)
        Me.btnImport.TabIndex = 1
        Me.btnImport.Text = "Upload"
        Me.btnImport.UseVisualStyleBackColor = False
        '
        'btnExport
        '
        Me.btnExport.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnExport.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnExport.Font = New System.Drawing.Font("Tahoma", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExport.ForeColor = System.Drawing.Color.White
        Me.btnExport.Location = New System.Drawing.Point(30, 167)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(288, 36)
        Me.btnExport.TabIndex = 0
        Me.btnExport.Text = "Download"
        Me.btnExport.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(30, 113)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(114, 30)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Main Data"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(381, 113)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(132, 30)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Branch Data"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Coral
        Me.Label3.Location = New System.Drawing.Point(22, 11)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(99, 47)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "DEM"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Microsoft Tai Le", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(27, 58)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(125, 16)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Data Export && iMports"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Maroon
        Me.Label5.Location = New System.Drawing.Point(114, 21)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(68, 29)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Apps"
        '
        'lblBranchLoading
        '
        Me.lblBranchLoading.BackColor = System.Drawing.Color.Transparent
        Me.lblBranchLoading.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblBranchLoading.Location = New System.Drawing.Point(376, 235)
        Me.lblBranchLoading.Name = "lblBranchLoading"
        Me.lblBranchLoading.Size = New System.Drawing.Size(287, 23)
        Me.lblBranchLoading.TabIndex = 10
        Me.lblBranchLoading.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pbBranchLoading
        '
        Me.pbBranchLoading.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.pbBranchLoading.Location = New System.Drawing.Point(376, 209)
        Me.pbBranchLoading.Name = "pbBranchLoading"
        Me.pbBranchLoading.Size = New System.Drawing.Size(287, 23)
        Me.pbBranchLoading.Step = 1
        Me.pbBranchLoading.TabIndex = 9
        '
        'lblClose
        '
        Me.lblClose.AutoSize = True
        Me.lblClose.BackColor = System.Drawing.Color.Transparent
        Me.lblClose.Font = New System.Drawing.Font("Verdana", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClose.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline
        Me.lblClose.LinkColor = System.Drawing.Color.Black
        Me.lblClose.Location = New System.Drawing.Point(664, 11)
        Me.lblClose.Name = "lblClose"
        Me.lblClose.Size = New System.Drawing.Size(28, 25)
        Me.lblClose.TabIndex = 11
        Me.lblClose.TabStop = True
        Me.lblClose.Text = "X"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("Microsoft Tai Le", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(104, 402)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(171, 16)
        Me.Label8.TabIndex = 29
        Me.Label8.Text = "Developed By ; Eric G. Canubas"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Font = New System.Drawing.Font("Microsoft Tai Le", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(104, 386)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(193, 16)
        Me.Label7.TabIndex = 28
        Me.Label7.Text = "Southwood Mindanao Corporation"
        '
        'FrmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.BackgroundImage = Global.DEM.My.Resources.Resources.orange
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(704, 447)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.lblClose)
        Me.Controls.Add(Me.lblBranchLoading)
        Me.Controls.Add(Me.pbBranchLoading)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblMainLoading)
        Me.Controls.Add(Me.btnExport)
        Me.Controls.Add(Me.pbMainLoading)
        Me.Controls.Add(Me.btnImport)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "FrmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "DEM App"
        Me.TransparencyKey = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnImport As Button
    Friend WithEvents btnExport As Button
    Friend WithEvents pbMainLoading As ProgressBar
    Friend WithEvents lblMainLoading As Label
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents lblBranchLoading As Label
    Friend WithEvents pbBranchLoading As ProgressBar
    Friend WithEvents lblClose As LinkLabel
    Friend WithEvents Label8 As Label
    Friend WithEvents Label7 As Label
End Class
