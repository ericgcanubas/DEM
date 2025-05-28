Public Class FrmMain
    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Dim DBNAME As String = CreateSmallDatabase()

        If DBNAME <> "" Then
            btnExport.Enabled = False
            Dim str As String = getConString(DBNAME)
            conn = New ADODB.Connection()
            conn.Open(str)

            CreateTable_tbl_PCPOS_Cashiers()
            Collect_tbl_PCPOS_Cashiers(pbLoading, lblLoading)

            CreateTable_tbl_bank()
            Collect_tbl_Bank(pbLoading, lblLoading)

            CreateTable_tbl_banks()
            Collect_tbl_Banks(pbLoading, lblLoading)

            CreateTable_tbl_Bank_Terms()
            Collect_tbl_Bank_Terms(pbLoading, lblLoading)

            ' always last
            CreateTable_tbl_ItemsForPLU()
            Collect_tbl_ItemsForPLU(pbLoading, lblLoading)

            btnExport.Enabled = True
            MessageBox.Show("Successfully Export")
        End If

        conn.Close()
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
        MessageBox.Show("o")
    End Sub
End Class