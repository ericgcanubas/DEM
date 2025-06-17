Imports System.Runtime.InteropServices
Imports System.IO
Public Class FrmBranch

    Dim GetReference As Double

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
    Private Sub EnableControl(e As Boolean)
        dtpDate.Enabled = e
        btnDownload.Enabled = e
        btnUpload.Enabled = e
        lblClose.Enabled = e

    End Sub
    Private Sub btnDownload_Click(sender As Object, e As EventArgs) Handles btnDownload.Click



        ' Download Branch Data
        Dim saveFileDialog As New SaveFileDialog()

        saveFileDialog.Filter = ""
        saveFileDialog.Title = "Save data"
        saveFileDialog.DefaultExt = ""
        Dim refFile As String = DateTime.Now.ToString("yyyyMMddHHmmss").ToLower()
        saveFileDialog.FileName = $"{gbl_Counter}_" & refFile

        If saveFileDialog.ShowDialog() = DialogResult.OK Then

            GL_EXPORT_PATH = saveFileDialog.FileName
            Dim DBNAME As String = CreateData()

            If DBNAME <> "" Then
                EnableControl(False)
                Dim str As String = getConString(DBNAME)
                ConnLocal = New ADODB.Connection()
                ConnLocal.ConnectionTimeout = 30
                ConnLocal.Open(str)

                Local_CreateTable_tbl_info(dtpDate.Value, refFile, gbl_Counter)

                ' vplus
                Branch_CreateTable_tbl_VPlus_Codes(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_VPlus_Codes_Validity(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_VPlus_Purchases_Points(pbBranchLoading, lblBranchLoading, dtpDate.Value)

                Branch_CreateTable_tbl_PS_GT(pbBranchLoading, lblBranchLoading)
                Branch_CreateTable_tbl_PS_GT_ZZ(pbBranchLoading, lblBranchLoading)

                'JOURNAL
                Branch_CreateTable_tbl_PS_E_Journal(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_PS_E_Journal_Detail(pbBranchLoading, lblBranchLoading, dtpDate.Value)


                'ADJUSTMENT JOURNAL
                Branch_CreateTable_tbl_PS_GT_Adjustment_EJournal(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_PS_GT_Adjustment_EJournal_Detail(pbBranchLoading, lblBranchLoading, dtpDate.Value)

                Branch_CreateTable_tbl_PS_EmployeeATD(pbBranchLoading, lblBranchLoading, dtpDate.Value)

                ' GIFT CERT
                Branch_CreateTable_tbl_GiftCert_List(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_GiftCert_Payment(pbBranchLoading, lblBranchLoading, dtpDate.Value)


                Branch_CreateTable_tbl_PS(pbBranchLoading, lblMainLoading, dtpDate.Value)
                Branch_CreateTable_tbl_PS_Tmp(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_PS_ItemsSold_Tmp(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_PS_ItemsSold_Voided(pbBranchLoading, lblBranchLoading, dtpDate.Value)

                Branch_CreateTable_tbl_PS_MiscPay(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_PS_MiscPay_Tmp(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_PS_MiscPay_Voided(pbBranchLoading, lblBranchLoading, dtpDate.Value)

                Branch_CreateTable_tbl_PaidOutTransactions(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_PaidOutTransactions_Det(pbBranchLoading, lblBranchLoading, dtpDate.Value)

                ' CREDIT MEMO
                Branch_CreateTable_tbl_CreditMemo(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_CreditMemo_CashRefund_Payment(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                Branch_CreateTable_tbl_CreditMemo_Payment(pbBranchLoading, lblBranchLoading, dtpDate.Value)
                ' HOME CREDIT
                Branch_CreateTable_tbl_HomeCredit_DeliveryAdvice(pbBranchLoading, lblBranchLoading, dtpDate.Value)

                SetLog(False)
                lblBranchLoading.Text = ""
                pbBranchLoading.Value = 0
                RefreshLog()
                ConnLocal.Close()
                ConnLocal = Nothing
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
                    GetBranchInfo()

                    If GetReference >= MainImportReference Then
                        MessageBox.Show("Invalid Upload. file already uploaded ")
                        ConnLocal.Close()
                        Exit Sub
                    End If

                    EnableControl(False)

                    Insert_Collect_tbl_CreditMemo(pbMainLoading, lblMainLoading)
                    Insert_Collect_tbl_HomeCredit_DeliveryAdvice(pbMainLoading, lblMainLoading)

                    Insert_Collect_tbl_PS_MiscPay_Tmp(pbMainLoading, lblMainLoading)
                    Insert_Collect_tbl_PS_MiscPay(pbMainLoading, lblMainLoading)

                    Insert_tbl_PS_GT(pbMainLoading, lblMainLoading)
                    Insert_tbl_PS_GT_ZZ(pbMainLoading, lblMainLoading)

                    Insert_tbl_PS_GT_Adjustment_EJournal(pbMainLoading, lblMainLoading)
                    Insert_tbl_PS_GT_Adjustment_EJournal_Detail(pbMainLoading, lblMainLoading)

                    Insert_tbl_PS_E_Journal(pbMainLoading, lblMainLoading)
                    Insert_tbl_PS_E_Journal_Detail(pbMainLoading, lblMainLoading)

                    Insert_tbl_PCPOS_Cashiers(pbMainLoading, lblMainLoading)
                    Insert_tbl_PCPOS_Cashiers_Changes(pbMainLoading, lblMainLoading)

                    Insert_tbl_Items_Changes(pbMainLoading, lblMainLoading)
                    Insert_tbl_ItemsForPLU_For_Effect(pbMainLoading, lblMainLoading)

                    Insert_tbl_Items(pbMainLoading, lblMainLoading)
                    Insert_tbl_ItemsForPLU(pbMainLoading, lblMainLoading)

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

                    Insert_tbl_Concession_PCR(pbMainLoading, lblMainLoading)
                    Insert_tbl_Concession_PCR_Det(pbMainLoading, lblMainLoading)
                    Insert_tbl_Concession_PCR_Effectivity(pbMainLoading, lblMainLoading)

                    Insert_tbl_GiftCert_Changes(pbMainLoading, lblMainLoading)
                    Insert_tbl_PS_Upload_Utility(pbMainLoading, lblMainLoading)

                    Insert_tbl_VPlus_Summary(pbMainLoading, lblMainLoading)
                    Insert_tbl_VPlus_Codes_For_Offline(pbMainLoading, lblMainLoading)

                    Insert_tbl_VPlus_App(pbMainLoading, lblMainLoading)
                    Insert_tbl_RetrieveHistoryForLocal(pbMainLoading, lblMainLoading)

                    Insert_tbl_PaidOutDenominations(pbMainLoading, lblMainLoading)
                    Insert_tbl_PaidOutTransactions(pbMainLoading, lblMainLoading)

                    Insert_Collect_tbl_PS_GT_History(pbMainLoading, lblMainLoading)
                    Insert_Collect_tbl_PS_GT_Zero_Out(pbMainLoading, lblMainLoading)

                    lblMainLoading.Text = ""
                    pbMainLoading.Value = 0

                    ConnLocal.Close()
                    SetLog(True)
                    RefreshLog()
                    EnableControl(True)
                    UpdateBranchInfo()
                    MessageBox.Show("Successfully Main Data Upload", "Upload Message", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Else
                    MessageBox.Show("Main data not found", "Upload Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

            Catch ex As Exception
                MessageBox.Show("Error uploading file: " & ex.Message)
            End Try
            EnableControl(True)
        End If
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

        Dim str As String = CreateDBBranch()

        ConnTemp = New ADODB.Connection()
        ConnTemp.ConnectionTimeout = 30
        ConnTemp.Open(str)
        Maketbl_UploadLog()
        MakeBranchInfo()

        EnableControl(True)
    End Sub

    Private Sub Maketbl_UploadLog()

        Try
            Dim createTableSql As String = "CREATE TABLE tbl_upload_log (
                                                [Counter] TEXT(5) PRIMARY KEY,
                                                DateUpload DATETIME,        
                                                [Reference] TEXT(15));"

            ConnTemp.Execute(createTableSql)
        Catch ex As Exception

        End Try



    End Sub

    Private Sub MakeBranchInfo()
        If gbl_Counter <> "" Then
            Try
                rs = New ADODB.Recordset
                rs.Open($"SELECT * FROM [tbl_upload_log] WHERE [Counter] = '{gbl_Counter}'", ConnTemp, ADODB.CursorTypeEnum.adOpenStatic)
                If rs.RecordCount = 0 Then
                    ConnTemp.Execute($"INSERT INTO [tbl_upload_log] ([Counter]) VALUES('{gbl_Counter}')  ")
                    MessageBox.Show($"New Branch Info {gbl_Counter}", "Branch Info")
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Branch Info")
            End Try

        End If
    End Sub


    Private Sub UpdateBranchInfo()
        Dim rx As New ADODB.Recordset
        rx.Open($"SELECT * FROM [tbl_upload_log] WHERE [Counter] = '{gbl_Counter}'", ConnTemp, ADODB.CursorTypeEnum.adOpenStatic)
        If rx.RecordCount <> 0 Then
            ConnTemp.Execute($"UPDATE [tbl_upload_log] SET DateUpload = '{Date.Now()}', [Reference] = '{MainImportReference}'  WHERE [Counter] = '{gbl_Counter}'")
        End If

    End Sub
    Public Sub GetBranchInfo()
        Dim rx As New ADODB.Recordset
        rx.Open($"SELECT * FROM [tbl_upload_log] WHERE [Counter] = '{gbl_Counter}'", ConnTemp, ADODB.CursorTypeEnum.adOpenStatic)
        If rx.RecordCount <> 0 Then
            GetReference = Val(rx.Fields("Reference").Value.ToString())
        Else
            GetReference = 0
        End If
    End Sub
End Class