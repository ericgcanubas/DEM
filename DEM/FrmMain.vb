Imports System.IO
Imports System.Runtime.InteropServices
Public Class FrmMain


    <DllImport("user32.dll")>
    Public Shared Function ReleaseCapture() As Boolean
    End Function

    <DllImport("user32.dll")>
    Public Shared Function SendMessage(hWnd As IntPtr, wMsg As Integer, wParam As Integer, lParam As Integer) As Integer
    End Function

    Const WM_NCLBUTTONDOWN As Integer = &HA1
    Const HTCAPTION As Integer = 2
    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        gbl_AdjustmentOnly = chkAdjustment.Checked

        saveIt()

    End Sub
    Private Sub saveIt()


        Dim saveFileDialog As New SaveFileDialog()

        ' Optional: Set filters and default settings
        ' Set filter for .mdb files
        saveFileDialog.Filter = ""
        saveFileDialog.Title = "Save data"
        saveFileDialog.DefaultExt = ""
        Dim ref As String = DateTime.Now.ToString("yyyyMMddHHmmss").ToLower()
        saveFileDialog.FileName = "main" & ref

        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            EnableControl(False)
            ' Get the selected file path
            GL_EXPORT_PATH = saveFileDialog.FileName
            Dim DBNAME As String = CreateData()
            If DBNAME <> "" Then
                btnExport.Enabled = False

                Dim str As String = getConString(DBNAME)
                ConnLocal = New ADODB.Connection()
                ConnLocal.ConnectionTimeout = 30
                ConnLocal.Open(str)

                Local_CreateTable_tbl_info(Now.Date, ref, "Main")

                gbl_DownloadType = Val(GetParameter("GenerateType"))
                NItemOnly = Val(GetParameter("ItemNotInclude"))



                CreateTable_tbl_banks(pbMainLoading, lblMainLoading)
                CreateTable_tbl_Banks_Changes(pbMainLoading, lblMainLoading)

                CreateTable_tbl_bank(pbMainLoading, lblMainLoading)
                CreateTable_tbl_Bank_Terms(pbMainLoading, lblMainLoading)
                CreateTable_tbl_Bank_Changes(pbMainLoading, lblMainLoading)

                CreateTable_tbl_QRPay_Type(pbMainLoading, lblMainLoading)

                CreateTable_tbl_GiftCert_List(pbMainLoading, lblMainLoading)
                CreateTable_tbl_GiftCert_Changes(pbMainLoading, lblMainLoading)

                CreateTable_tbl_VPlus_Codes(pbMainLoading, lblMainLoading)
                CreateTable_tbl_VPlus_Codes_Validity(pbMainLoading, lblMainLoading)
                CreateTable_tbl_VPlus_Codes_Changes(pbMainLoading, lblMainLoading)
                CreateTable_tbl_VPlus_Summary(pbMainLoading, lblMainLoading)
                CreateTable_tbl_VPlus_Codes_For_Offline(pbMainLoading, lblMainLoading)
                CreateTable_tbl_VPlus_App(pbMainLoading, lblMainLoading)


                CreateTable_tbl_PS_GT_Adjustment_EJournal(pbMainLoading, lblMainLoading)
                CreateTable_tbl_PS_GT_Adjustment_EJournal_Detail(pbMainLoading, lblMainLoading)

                CreateTable_tbl_PS_E_Journal(pbMainLoading, lblMainLoading)
                CreateTable_tbl_PS_E_Journal_Detail(pbMainLoading, lblMainLoading)
                CreateTable_tbl_PS_GT_History(pbMainLoading, lblMainLoading)
                CreateTable_tbl_PS_GT_Zero_Out(pbMainLoading, lblMainLoading)

                CreateTable_tbl_PS_GT(pbMainLoading, lblMainLoading)

                CreateTable_tbl_PS_GT_ZZ(pbMainLoading, lblMainLoading)

                CreateTable_tbl_PS_Upload_Utility(pbMainLoading, lblMainLoading)

                CreateTable_tbl_PCPOS_Cashiers(pbMainLoading, lblMainLoading)
                CreateTable_tbl_PCPOS_Cashiers_Changes(pbMainLoading, lblMainLoading)


                CreateTable_tbl_Concession_PCR(pbMainLoading, lblMainLoading)
                CreateTable_tbl_Concession_PCR_Det(pbMainLoading, lblMainLoading)
                CreateTable_tbl_Concession_PCR_Effectivity(pbMainLoading, lblMainLoading)

                CreateTable_tbl_RetrieveHistoryForLocal(pbMainLoading, lblMainLoading)

                CreateTable_tbl_Items(pbMainLoading, lblMainLoading)
                CreateTable_tbl_Items_Change(pbMainLoading, lblMainLoading)

                CreateTable_tbl_ItemsForPLU(pbMainLoading, lblMainLoading)
                CreateTable_tbl_ItemsForPLU_For_Effect(pbMainLoading, lblMainLoading)

                CreateTable_tbl_PaidOutDenominations(pbMainLoading, lblMainLoading)
                CreateTable_tbl_PaidOutTransactions(pbMainLoading, lblMainLoading)

                SetLog(False)
                lblMainLoading.Text = ""
                btnExport.Enabled = True
                RefreshLog()

                EnableControl(True)
                Dim result As DialogResult = MessageBox.Show(
                "File saved successfully at:" & vbCrLf & GL_EXPORT_PATH & vbCrLf & vbCrLf &
                "Do you want to open the location?",
                "File Saved",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            )
                If result = DialogResult.Yes Then
                    Process.Start("explorer.exe", "/select,""" & GL_EXPORT_PATH & """")
                End If
            End If

            Try
                ConnLocal.Close()
            Catch ex As Exception

            End Try

            EnableControl(True)
        End If

    End Sub
    Private Sub EnableControl(e As Boolean)
        picOpenMain.Enabled = e
        btnExport.Enabled = e
        btnImport.Enabled = e
        lblClose.Enabled = e

    End Sub
    Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        gbl_Server = GetSetting("DEM", "MODE", "SERVER")
        gbl_Database = GetSetting("DEM", "MODE", "DATABASE")
        lblSERVER.Text = $"SN:{gbl_Server}"
        lblDATABASE.Text = $"DB:{gbl_Database}"
        getConnection()

        Dim str As String = CreateDBMain()

        ConnTemp = New ADODB.Connection()
        ConnTemp.ConnectionTimeout = 30
        ConnTemp.Open(str)
        Maketbl_COUNTER()
        Maketbl_PARAMETER()
        gbl_DownloadType = Val(GetParameter("GenerateType"))
        NItemOnly = Val(GetParameter("ItemNotInclude"))
        RefreshLog()
    End Sub
    Private Sub RefreshLog()
        lblLogDownload.Text = $"Last Download On : { GetLog(False)}"
        lblLogUpload.Text = $"Last Upload On : { GetLog(True)}"
    End Sub
    Private Sub Maketbl_COUNTER()

        Try
            Dim createTableSql As String = "CREATE TABLE tbl_counter_list (
                                                [Counter] TEXT(5) PRIMARY KEY,
                                                DateUpload DATETIME,        
                                                [Reference] TEXT(15)
                                        );"

            ConnTemp.Execute(createTableSql)
        Catch ex As Exception

        End Try

    End Sub
    Private Sub Maketbl_PARAMETER()

        Try
            Dim createTableSql As String = "CREATE TABLE tbl_parameter (
                                                ParameterName TEXT(15) PRIMARY KEY,
                                                ParameterValue TEXT(50)                                                   
                                        );"
            ConnTemp.Execute(createTableSql)
        Catch ex As Exception

        End Try

    End Sub


    Private Sub pbLoading_Click(sender As Object, e As EventArgs) Handles pbMainLoading.Click

    End Sub

    Private Sub pbLoading_TextChanged(sender As Object, e As EventArgs) Handles pbMainLoading.TextChanged

    End Sub

    Private Sub pbLoading_RegionChanged(sender As Object, e As EventArgs) Handles pbMainLoading.RegionChanged

    End Sub

    Private Sub FrmMain_Closed(sender As Object, e As EventArgs) Handles Me.Closed


        Try
            ConnLocal.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub FrmMain_MouseDown(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDown
        If e.Button = MouseButtons.Left Then
            ReleaseCapture()
            SendMessage(Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0)
        End If
    End Sub

    Private Sub lblClose_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lblClose.LinkClicked
        End
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        branchInsert()
    End Sub

    Private Sub branchInsert()
        Dim ofd As New OpenFileDialog()
        ofd.Title = "Select a file to upload"
        ofd.Filter = "All Files (*.*)|*.*"

        ' Show the dialog and check if the user selected a file
        If ofd.ShowDialog() = DialogResult.OK Then
            Try
                Dim sourceFilePath As String = ofd.FileName
                Dim fileName As String = Path.GetFileName(sourceFilePath) ' file name

                Dim str As String = getConnectionString(sourceFilePath)
                ConnLocal = New ADODB.Connection()
                ConnLocal.ConnectionTimeout = 30
                ConnLocal.Open(str)
                If GetBranchInfo() = True Then
                    EnableControl(False)
                    Branch_Insert_tbl_GiftCert_List(pbBranchLoading, lblBranchLoading)
                    Branch_Insert_tbl_VPlus_Codes(pbBranchLoading, lblBranchLoading)
                    Branch_Insert_tbl_VPlus_Codes_Validity(pbBranchLoading, lblBranchLoading)
                    Branch_Insert_tbl_PS_GT(pbBranchLoading, lblBranchLoading)
                    Branch_Insert_tbl_PS_GT_ZZ(pbBranchLoading, lblBranchLoading)

                    Branch_Insert_tbl_PS_E_Journal(pbBranchLoading, lblBranchLoading)
                    Branch_Insert_tbl_PS_E_Journal_Detail(pbBranchLoading, lblBranchLoading)

                    Branch_Insert_tbl_PS_GT_Adjustment_EJournal(pbBranchLoading, lblBranchLoading)
                    Branch_Insert_tbl_PS_GT_Adjustment_EJournal_Detail(pbBranchLoading, lblBranchLoading)

                    Branch_Insert_tbl_PS_EmployeeATD(pbBranchLoading, lblBranchLoading)
                    Branch_Insert_tbl_GiftCert_Payment(pbBranchLoading, lblBranchLoading)
                    Branch_Insert_tbl_VPlus_Purchases_Points(pbBranchLoading, lblBranchLoading)

                    Branch_Insert_tbl_PS(pbBranchLoading, lblBranchLoading)
                    Branch_Insert_tbl_PS_Tmp(pbBranchLoading, lblBranchLoading)
                    Branch_Insert_tbl_PS_ItemsSold_Tmp(pbBranchLoading, lblBranchLoading)
                    Branch_Insert_tbl_PS_ItemsSold_Voided(pbBranchLoading, lblBranchLoading)
                    Branch_Insert_tbl_PS_MiscPay_Tmp(pbBranchLoading, lblBranchLoading)
                    Branch_Insert_tbl_PS_MiscPay_Voided(pbBranchLoading, lblBranchLoading)
                    Branch_Insert_tbl_PaidOutTransactions(pbBranchLoading, lblBranchLoading)
                    Branch_Insert__tbl_ItemTransactions(pbBranchLoading, lblBranchLoading)
                End If
                ConnLocal.Close()
                MessageBox.Show("Successfully Branch Data Upload", "Upload Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
                pbBranchLoading.Value = 0
                lblBranchLoading.Text = ""
                SetLog(True)
                RefreshLog()
                EnableControl(True)
            Catch ex As Exception
                MessageBox.Show("Error uploading file: " & ex.Message)
            End Try
            EnableControl(True)
        End If
    End Sub

    Private Sub picOpenMain_Click(sender As Object, e As EventArgs) Handles picOpenMain.Click
        FrmMainSetup.ShowDialog()
        FrmMainSetup = Nothing
    End Sub
End Class