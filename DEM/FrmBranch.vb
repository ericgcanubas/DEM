Imports System.Runtime.InteropServices
Imports System.IO
Public Class FrmBranch
    <DllImport("user32.dll")>
    Public Shared Function ReleaseCapture() As Boolean
    End Function

    <DllImport("user32.dll")>
    Public Shared Function SendMessage(hWnd As IntPtr, wMsg As Integer, wParam As Integer, lParam As Integer) As Integer
    End Function

    Const WM_NCLBUTTONDOWN As Integer = &HA1
    Const HTCAPTION As Integer = 2
    Private Sub lblClose_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lblClose.LinkClicked
        End
    End Sub

    Private Sub FrmBranch_MouseDown(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDown
        If e.Button = MouseButtons.Left Then
            ReleaseCapture()
            SendMessage(Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0)
        End If
    End Sub

    Private Sub btnDownload_Click(sender As Object, e As EventArgs) Handles btnDownload.Click


        Dim saveFileDialog As New SaveFileDialog()

        ' Optional: Set filters and default settings
        ' Set filter for .mdb files
        saveFileDialog.Filter = ""
        saveFileDialog.Title = "Save data"
        saveFileDialog.DefaultExt = ""
        Dim refFile As String = DateTime.Now.ToString("yyyyMMddHHmmss").ToLower()
        saveFileDialog.FileName = $"{gbl_Counter}_" & refFile

        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            ' Get the selected file path
            GL_EXPORT_PATH = saveFileDialog.FileName
            Dim DBNAME As String = CreateData()

            If DBNAME <> "" Then

                Dim str As String = getConString(DBNAME)
                ConnLocal = New ADODB.Connection()
                ConnLocal.ConnectionTimeout = 30
                ConnLocal.Open(str)

                Local_CreateTable_tbl_info(dtpDate.Value, refFile, gbl_Counter)

                Branch_CreateTable_tbl_VPlus_Codes(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_VPlus_Codes_Validity(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_VPlus_Purchases_Points(pbBranchLoading, lblBranchLoading, dtpDate.Value)

                Branch_CreateTable_tbl_PS_GT(pbBranchLoading, lblBranchLoading)
                Branch_CreateTable_tbl_PS_GT_ZZ(pbBranchLoading, lblBranchLoading)

                Branch_CreateTable_tbl_PS_E_Journal(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_PS_E_Journal_Detail(pbBranchLoading, lblBranchLoading, dtpDate.Value)

                Branch_CreateTable_tbl_PS_EmployeeATD(pbBranchLoading, lblBranchLoading, dtpDate.Value)

                Branch_CreateTable_tbl_GiftCert_List(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_GiftCert_Payment(pbBranchLoading, lblBranchLoading, dtpDate.Value)

                Branch_CreateTable_tbl_PS_Tmp(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_PS_ItemsSold_Tmp(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_PS_ItemsSold_Voided(pbBranchLoading, lblBranchLoading, dtpDate.Value)

                Branch_CreateTable_tbl_PS_MiscPay_Tmp(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_PS_MiscPay_Voided(pbBranchLoading, lblBranchLoading, dtpDate.Value)

                Branch_CreateTable_tbl_PaidOutTransactions(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_PaidOutTransactions_Det(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                SetLog(False)
                lblBranchLoading.Text = ""
                pbBranchLoading.Value = 0
                RefreshLog()
                ConnLocal.Close()
                ConnLocal = Nothing

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
        End If
    End Sub

    Private Sub btnUpload_Click(sender As Object, e As EventArgs) Handles btnUpload.Click
        ' Create and configure the OpenFileDialog
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

                If GetMainInfo() = True Then
                    SaveIt()

                    ConnLocal.Close()
                    SetLog(True)
                    RefreshLog()
                    MessageBox.Show("Successfully Main Data Upload", "Upload Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Main data not found", "Upload Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If



            Catch ex As Exception
                MessageBox.Show("Error uploading file: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub SaveIt()
        Insert_tbl_PaidOutDenominations(pbMainLoading, lblMainLoading)
        Insert_tbl_PaidOutTransactions(pbMainLoading, lblMainLoading)

        Insert_tbl_Bank(pbMainLoading, lblMainLoading)
        Insert_tbl_Banks(pbMainLoading, lblMainLoading)
        Insert_tbl_Banks_Changes(pbMainLoading, lblMainLoading)
        Insert_tbl_Bank_Changes(pbMainLoading, lblMainLoading)
        Insert_tbl_Bank_Terms(pbMainLoading, lblMainLoading)

        Insert_tbl_QRPay_Type(pbMainLoading, lblMainLoading)
        Insert_tbl_GiftCert_List(pbMainLoading, lblMainLoading)

        Insert_tbl_VPlus_Codes(pbMainLoading, lblMainLoading)
        Insert_tbl_VPlus_Codes_Validity(pbMainLoading, lblMainLoading)
        Insert_tbl_VPlus_Codes_Changes(pbMainLoading, lblMainLoading)

        Insert_tbl_PCPOS_Cashiers_Changes(pbMainLoading, lblMainLoading)

        Insert_tbl_Items_Changes(pbMainLoading, lblMainLoading)
        Insert_tbl_ItemsForPLU_For_Effect(pbMainLoading, lblMainLoading)
        Insert_tbl_Items(pbMainLoading, lblMainLoading)


        Insert_tbl_Concession_PCR(pbMainLoading, lblMainLoading)
        Insert_tbl_Concession_PCR_Det(pbMainLoading, lblMainLoading)
        Insert_tbl_Concession_PCR_Effectivity(pbMainLoading, lblMainLoading)

        Insert_tbl_GiftCert_Changes(pbMainLoading, lblMainLoading)
        Insert_tbl_PS_Upload_Utility(pbMainLoading, lblMainLoading)

        Insert_tbl_VPlus_Summary(pbMainLoading, lblMainLoading)
        Insert_tbl_VPlus_Codes_For_Offline(pbMainLoading, lblMainLoading)
        Insert_tbl_VPlus_App(pbMainLoading, lblMainLoading)
        Insert_tbl_RetrieveHistoryForLocal(pbMainLoading, lblMainLoading)

        Insert_tbl_PS_GT(pbMainLoading, lblMainLoading)
        Insert_tbl_PS_GT_ZZ(pbMainLoading, lblMainLoading)

        Insert_tbl_PS_E_Journal(pbMainLoading, lblMainLoading)
        Insert_tbl_PS_E_Journal_Detail(pbMainLoading, lblMainLoading)

        Insert_tbl_PS_GT_Adjustment_EJournal(pbMainLoading, lblMainLoading)
        Insert_tbl_PS_GT_Adjustment_EJournal_Detail(pbMainLoading, lblMainLoading)


        Insert_tbl_PCPOS_Cashiers(pbMainLoading, lblMainLoading)
        Insert_tbl_ItemsForPLU(pbMainLoading, lblMainLoading)


        lblMainLoading.Text = ""
        pbMainLoading.Value = 0

    End Sub
    Private Sub RefreshLog()
        lblLogDownload.Text = $"Last Download On : { GetLog(False)}"
        lblLogUpload.Text = $"Last Upload On : { GetLog(True)}"
    End Sub
    Private Sub FrmBranch_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        gbl_Server = GetSetting("DEM", "MODE", "SERVER")
        gbl_Database = GetSetting("DEM", "MODE", "DATABASE")
        gbl_Counter = GetSetting("DEM", "MODE", "COUNTER")
        lblCOUNTER.Text = $"COUNTER : {gbl_Counter}"
        getConnection()
        RefreshLog()
    End Sub
End Class