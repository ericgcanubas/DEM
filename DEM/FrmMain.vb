Public Class FrmMain
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
            Dim DBNAME As String = CreateSmallDatabase()

            If DBNAME <> "" Then
                btnExport.Enabled = False

                Dim str As String = getConString(DBNAME)
                conn = New ADODB.Connection()
                conn.Open(str)


                CreateTable_tbl_banks(pbLoading, lblLoading)
                CreateTable_tbl_Banks_Changes(pbLoading, lblLoading)

                CreateTable_tbl_bank(pbLoading, lblLoading)
                CreateTable_tbl_Bank_Terms(pbLoading, lblLoading)
                CreateTable_tbl_Bank_Changes(pbLoading, lblLoading)

                CreateTable_tbl_QRPay_Type(pbLoading, lblLoading)

                CreateTable_tbl_GiftCert_List(pbLoading, lblLoading)
                CreateTable_tbl_GiftCert_Changes(pbLoading, lblLoading)

                CreateTable_tbl_VPlus_Codes(pbLoading, lblLoading)
                CreateTable_tbl_VPlus_Codes_Validity(pbLoading, lblLoading)


                CreateTable_tbl_PCPOS_Cashiers(pbLoading, lblLoading)
                CreateTable_tbl_PCPOS_Cashiers_Changes(pbLoading, lblLoading)

                CreateTable_tbl_Concession_PCR_Effectivity(pbLoading, lblLoading)
                CreateTable_tbl_Concession_PCR(pbLoading, lblLoading)
                CreateTable_tbl_Concession_PCR_Det(pbLoading, lblLoading)

                CreateTable_tbl_Items(pbLoading, lblLoading)
                CreateTable_tbl_Items_Change(pbLoading, lblLoading)
                CreateTable_tbl_ItemsForPLU(pbLoading, lblLoading)
                CreateTable_tbl_ItemsForPLU_For_Effect(pbLoading, lblLoading)

                lblLoading.Text = ""
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
                conn.Close()
            Catch ex As Exception

            End Try


        End If

    End Sub
    Private Sub FrmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        gbl_Server = GetSetting("SYNCRONIZER", "MODE", "SERVER")
        gbl_Database = GetSetting("SYNCRONIZER", "MODE", "DATABASE")
        getConnection()
    End Sub

    Private Sub pbLoading_Click(sender As Object, e As EventArgs) Handles pbLoading.Click

    End Sub

    Private Sub pbLoading_TextChanged(sender As Object, e As EventArgs) Handles pbLoading.TextChanged

    End Sub

    Private Sub pbLoading_RegionChanged(sender As Object, e As EventArgs) Handles pbLoading.RegionChanged

    End Sub

    Private Sub FrmMain_Closed(sender As Object, e As EventArgs) Handles Me.Closed


        Try
            conn.Close()
        Catch ex As Exception

        End Try
    End Sub
End Class