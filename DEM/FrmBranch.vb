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
                SaveIt()

                ConnLocal.Close()
            Catch ex As Exception
                MessageBox.Show("Error uploading file: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub SaveIt()


        Insert_tbl_Bank(pbMainLoading, lblMainLoading)
        Insert_tbl_Banks(pbMainLoading, lblMainLoading)
        Insert_tbl_Banks_Changes(pbMainLoading, lblMainLoading)
        Insert_tbl_Bank_Changes(pbMainLoading, lblMainLoading)
        Insert_tbl_Bank_Terms(pbMainLoading, lblMainLoading)
        Insert_tbl_QRPay_Type(pbMainLoading, lblMainLoading)
        Insert_tbl_GiftCert_List(pbMainLoading, lblMainLoading)
        Insert_tbl_VPlus_Codes(pbMainLoading, lblMainLoading)
        Insert_tbl_VPlus_Codes_Validity(pbMainLoading, lblMainLoading)
        Insert_tbl_PCPOS_Cashiers_Changes(pbMainLoading, lblMainLoading)
        Insert_tbl_Items_Changes(pbMainLoading, lblMainLoading)
        Insert_tbl_ItemsForPLU_For_Effect(pbMainLoading, lblMainLoading)
        Insert_tbl_Items(pbMainLoading, lblMainLoading)

        Insert_tbl_PCPOS_Cashiers(pbMainLoading, lblMainLoading)
        Insert_tbl_ItemsForPLU(pbMainLoading, lblMainLoading)



    End Sub

    Private Sub FrmBranch_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        gbl_Server = GetSetting("SYNCRONIZER", "MODE", "SERVER")
        gbl_Database = GetSetting("SYNCRONIZER", "MODE", "DATABASE")
        getConnection()
    End Sub
End Class