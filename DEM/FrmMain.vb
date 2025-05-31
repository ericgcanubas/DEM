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

        saveIt()

    End Sub
    Private Sub saveIt()


        Dim saveFileDialog As New SaveFileDialog()

        ' Optional: Set filters and default settings
        ' Set filter for .mdb files
        saveFileDialog.Filter = ""
        saveFileDialog.Title = "Save data"
        saveFileDialog.DefaultExt = ""
        saveFileDialog.FileName = "main" & DateTime.Now.ToString("yyyyMMddHHmmss").ToLower() & ""

        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            ' Get the selected file path
            GL_EXPORT_PATH = saveFileDialog.FileName
            Dim DBNAME As String = CreateData()

            If DBNAME <> "" Then
                btnExport.Enabled = False

                Dim str As String = getConString(DBNAME)
                ConnLocal = New ADODB.Connection()
                ConnLocal.ConnectionTimeout = 30
                ConnLocal.Open(str)

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

                CreateTable_tbl_PS_GT_Adjustment_EJournal_Detail(pbMainLoading, lblMainLoading)
                CreateTable_tbl_PS_GT_Adjustment_EJournal(pbMainLoading, lblMainLoading)
                CreateTable_tbl_PS_E_Journal(pbMainLoading, lblMainLoading)
                CreateTable_tbl_PS_E_Journal_Detail(pbMainLoading, lblMainLoading)

                CreateTable_tbl_PS_GT(pbMainLoading, lblMainLoading)
                CreateTable_tbl_PS_GT_ZZ(pbMainLoading, lblMainLoading)

                CreateTable_tbl_PS_Upload_Utility(pbMainLoading, lblMainLoading)

                CreateTable_tbl_PCPOS_Cashiers(pbMainLoading, lblMainLoading)
                CreateTable_tbl_PCPOS_Cashiers_Changes(pbMainLoading, lblMainLoading)

                CreateTable_tbl_Concession_PCR_Effectivity(pbMainLoading, lblMainLoading)
                CreateTable_tbl_Concession_PCR(pbMainLoading, lblMainLoading)
                CreateTable_tbl_Concession_PCR_Det(pbMainLoading, lblMainLoading)

                CreateTable_tbl_RetrieveHistoryForLocal(pbMainLoading, lblMainLoading)

                CreateTable_tbl_Items(pbMainLoading, lblMainLoading)
                CreateTable_tbl_Items_Change(pbMainLoading, lblMainLoading)
                CreateTable_tbl_ItemsForPLU(pbMainLoading, lblMainLoading)
                CreateTable_tbl_ItemsForPLU_For_Effect(pbMainLoading, lblMainLoading)

                lblMainLoading.Text = ""
                btnExport.Enabled = True

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


        End If

    End Sub
    Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        gbl_Server = GetSetting("DEM", "MODE", "SERVER")
        gbl_Database = GetSetting("DEM", "MODE", "DATABASE")
        getConnection()
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
End Class