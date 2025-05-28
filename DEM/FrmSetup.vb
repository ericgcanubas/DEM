Public Class FrmSetup
    Private Sub FrmSetup_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim Type As String = GetSetting("SYNCRONIZER", "MODE", "TYPE")

        If Type <> "" Then
            If Type = False Then
                FrmBranch.Show()
            Else
                FrmMain.Show()
            End If

            Me.Close()
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        If rdBranch.Checked = False And rdMainOffice.Checked = False Then
            MessageBox.Show("Please select option")

        ElseIf (txtServer.Text.Length = 0) Then
            MessageBox.Show("Please enter server name")

        ElseIf (txtDatabase.Text.Length = 0) Then
            MessageBox.Show("Please enter database")
        Else
            If rdBranch.Checked = True Then
                SaveSetting("SYNCRONIZER", "MODE", "TYPE", "0")
            Else
                SaveSetting("SYNCRONIZER", "MODE", "TYPE", "1")
            End If

            SaveSetting("SYNCRONIZER", "MODE", "SERVER", txtServer.Text)
            SaveSetting("SYNCRONIZER", "MODE", "DATABASE", txtDatabase.Text)


            MessageBox.Show("Save please try run again")
            Application.Exit()
        End If
    End Sub
End Class