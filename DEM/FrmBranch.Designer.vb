﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmBranch
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmBranch))
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.lblBranchLoading = New System.Windows.Forms.Label()
        Me.btnUpload = New System.Windows.Forms.Button()
        Me.pbBranchLoading = New System.Windows.Forms.ProgressBar()
        Me.btnDownload = New System.Windows.Forms.Button()
        Me.SaveFileDialog2 = New System.Windows.Forms.SaveFileDialog()
        Me.lblClose = New System.Windows.Forms.LinkLabel()
        Me.lblMainLoading = New System.Windows.Forms.Label()
        Me.pbMainLoading = New System.Windows.Forms.ProgressBar()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dtpDate = New System.Windows.Forms.DateTimePicker()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.lblLogUpload = New System.Windows.Forms.Label()
        Me.lblLogDownload = New System.Windows.Forms.Label()
        Me.lblCOUNTER = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblBranchLoading
        '
        Me.lblBranchLoading.BackColor = System.Drawing.Color.Transparent
        Me.lblBranchLoading.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblBranchLoading.Location = New System.Drawing.Point(42, 249)
        Me.lblBranchLoading.Name = "lblBranchLoading"
        Me.lblBranchLoading.Size = New System.Drawing.Size(287, 23)
        Me.lblBranchLoading.TabIndex = 15
        Me.lblBranchLoading.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnUpload
        '
        Me.btnUpload.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnUpload.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnUpload.Font = New System.Drawing.Font("Tahoma", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpload.ForeColor = System.Drawing.Color.White
        Me.btnUpload.Location = New System.Drawing.Point(387, 181)
        Me.btnUpload.Name = "btnUpload"
        Me.btnUpload.Size = New System.Drawing.Size(288, 36)
        Me.btnUpload.TabIndex = 12
        Me.btnUpload.Text = "Upload"
        Me.btnUpload.UseVisualStyleBackColor = False
        '
        'pbBranchLoading
        '
        Me.pbBranchLoading.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.pbBranchLoading.Location = New System.Drawing.Point(42, 223)
        Me.pbBranchLoading.Name = "pbBranchLoading"
        Me.pbBranchLoading.Size = New System.Drawing.Size(287, 23)
        Me.pbBranchLoading.Step = 1
        Me.pbBranchLoading.TabIndex = 14
        '
        'btnDownload
        '
        Me.btnDownload.BackColor = System.Drawing.Color.Moccasin
        Me.btnDownload.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnDownload.Font = New System.Drawing.Font("Tahoma", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDownload.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnDownload.Location = New System.Drawing.Point(42, 181)
        Me.btnDownload.Name = "btnDownload"
        Me.btnDownload.Size = New System.Drawing.Size(288, 36)
        Me.btnDownload.TabIndex = 13
        Me.btnDownload.Text = "Download"
        Me.btnDownload.UseVisualStyleBackColor = False
        '
        'lblClose
        '
        Me.lblClose.AutoSize = True
        Me.lblClose.BackColor = System.Drawing.Color.Transparent
        Me.lblClose.Font = New System.Drawing.Font("Verdana", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClose.LinkBehavior = System.Windows.Forms.LinkBehavior.NeverUnderline
        Me.lblClose.LinkColor = System.Drawing.Color.Black
        Me.lblClose.Location = New System.Drawing.Point(652, 29)
        Me.lblClose.Name = "lblClose"
        Me.lblClose.Size = New System.Drawing.Size(28, 25)
        Me.lblClose.TabIndex = 23
        Me.lblClose.TabStop = True
        Me.lblClose.Text = "X"
        '
        'lblMainLoading
        '
        Me.lblMainLoading.BackColor = System.Drawing.Color.Transparent
        Me.lblMainLoading.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblMainLoading.Location = New System.Drawing.Point(387, 249)
        Me.lblMainLoading.Name = "lblMainLoading"
        Me.lblMainLoading.Size = New System.Drawing.Size(287, 23)
        Me.lblMainLoading.TabIndex = 22
        Me.lblMainLoading.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pbMainLoading
        '
        Me.pbMainLoading.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.pbMainLoading.Location = New System.Drawing.Point(387, 223)
        Me.pbMainLoading.Name = "pbMainLoading"
        Me.pbMainLoading.Size = New System.Drawing.Size(287, 23)
        Me.pbMainLoading.Step = 1
        Me.pbMainLoading.TabIndex = 21
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Maroon
        Me.Label5.Location = New System.Drawing.Point(120, 33)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(68, 29)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Apps"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Microsoft Tai Le", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(33, 70)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(159, 19)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Data Export && iMports"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Coral
        Me.Label3.Location = New System.Drawing.Point(28, 23)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(99, 47)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "DEM"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(48, 115)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(81, 30)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Branch"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(382, 115)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 30)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Main"
        '
        'dtpDate
        '
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDate.Location = New System.Drawing.Point(162, 152)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(112, 20)
        Me.dtpDate.TabIndex = 24
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Font = New System.Drawing.Font("Microsoft Tai Le", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(50, 152)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(106, 16)
        Me.Label7.TabIndex = 25
        Me.Label7.Text = "Date Transaction : "
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Font = New System.Drawing.Font("Microsoft Tai Le", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(103, 379)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(193, 16)
        Me.Label6.TabIndex = 26
        Me.Label6.Text = "Southwood Mindanao Corporation"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("Microsoft Tai Le", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(103, 395)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(171, 16)
        Me.Label8.TabIndex = 27
        Me.Label8.Text = "Developed By : Eric G. Canubas"
        '
        'lblLogUpload
        '
        Me.lblLogUpload.BackColor = System.Drawing.Color.Transparent
        Me.lblLogUpload.Font = New System.Drawing.Font("Microsoft Tai Le", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLogUpload.Location = New System.Drawing.Point(384, 272)
        Me.lblLogUpload.Name = "lblLogUpload"
        Me.lblLogUpload.Size = New System.Drawing.Size(290, 16)
        Me.lblLogUpload.TabIndex = 36
        Me.lblLogUpload.Text = "Last Upload :"
        '
        'lblLogDownload
        '
        Me.lblLogDownload.BackColor = System.Drawing.Color.Transparent
        Me.lblLogDownload.Font = New System.Drawing.Font("Microsoft Tai Le", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLogDownload.Location = New System.Drawing.Point(40, 272)
        Me.lblLogDownload.Name = "lblLogDownload"
        Me.lblLogDownload.Size = New System.Drawing.Size(290, 16)
        Me.lblLogDownload.TabIndex = 35
        Me.lblLogDownload.Text = "Last Download : "
        '
        'lblCOUNTER
        '
        Me.lblCOUNTER.AutoSize = True
        Me.lblCOUNTER.BackColor = System.Drawing.Color.Transparent
        Me.lblCOUNTER.Font = New System.Drawing.Font("Microsoft Tai Le", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCOUNTER.Location = New System.Drawing.Point(529, 38)
        Me.lblCOUNTER.Name = "lblCOUNTER"
        Me.lblCOUNTER.Size = New System.Drawing.Size(62, 16)
        Me.lblCOUNTER.TabIndex = 37
        Me.lblCOUNTER.Text = "COUTER : "
        '
        'FrmBranch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.BackgroundImage = Global.DEM.My.Resources.Resources.orange
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(704, 447)
        Me.Controls.Add(Me.lblCOUNTER)
        Me.Controls.Add(Me.lblLogUpload)
        Me.Controls.Add(Me.lblLogDownload)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.dtpDate)
        Me.Controls.Add(Me.lblBranchLoading)
        Me.Controls.Add(Me.btnUpload)
        Me.Controls.Add(Me.pbBranchLoading)
        Me.Controls.Add(Me.btnDownload)
        Me.Controls.Add(Me.lblClose)
        Me.Controls.Add(Me.lblMainLoading)
        Me.Controls.Add(Me.pbMainLoading)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmBranch"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Branch"
        Me.TransparencyKey = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents lblBranchLoading As Label
    Friend WithEvents btnUpload As Button
    Friend WithEvents pbBranchLoading As ProgressBar
    Friend WithEvents btnDownload As Button
    Friend WithEvents SaveFileDialog2 As SaveFileDialog
    Friend WithEvents lblClose As LinkLabel
    Friend WithEvents lblMainLoading As Label
    Friend WithEvents pbMainLoading As ProgressBar
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents dtpDate As DateTimePicker
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents lblLogUpload As Label
    Friend WithEvents lblLogDownload As Label
    Friend WithEvents lblCOUNTER As Label
End Class
