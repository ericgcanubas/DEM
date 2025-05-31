Imports System.Runtime.InteropServices
Public Class FrmSetup
    <DllImport("user32.dll")>
    Public Shared Function ReleaseCapture() As Boolean
    End Function

    <DllImport("user32.dll")>
    Public Shared Function SendMessage(hWnd As IntPtr, wMsg As Integer, wParam As Integer, lParam As Integer) As Integer
    End Function

    Const WM_NCLBUTTONDOWN As Integer = &HA1
    Const HTCAPTION As Integer = 2
    Private Sub FrmSetup_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim Type As String = GetSetting("DEM", "MODE", "TYPE")

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
                If (txtCounter.Text.Trim().Length = 0) Then
                    MessageBox.Show("Please enter Counter/POS Name")
                    Exit Sub
                End If
            End If
            If rdBranch.Checked = True Then
                SaveSetting("DEM", "MODE", "TYPE", "0")
            Else
                SaveSetting("DEM", "MODE", "TYPE", "1")
            End If

            SaveSetting("DEM", "MODE", "SERVER", txtServer.Text)
            SaveSetting("DEM", "MODE", "DATABASE", txtDatabase.Text)
            SaveSetting("DEM", "MODE", "COUNTER", txtCounter.Text)

            SaveSetting("DEM", "MODE", "UPLOAD_LOG", "")
            SaveSetting("DEM", "MODE", "DOWNLOAD_LOG", "")

            MessageBox.Show("Save please try run again")
            Application.Exit()
        End If
    End Sub

    Private Sub rdBranch_CheckedChanged(sender As Object, e As EventArgs) Handles rdBranch.CheckedChanged
        txtCounter.Enabled = True
    End Sub

    Private Sub rdMainOffice_CheckedChanged(sender As Object, e As EventArgs) Handles rdMainOffice.CheckedChanged
        txtCounter.Enabled = False
        txtCounter.Clear()

    End Sub

    Private Sub lblClose_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lblClose.LinkClicked
        End
    End Sub

    Private Sub FrmSetup_MouseDown(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDown
        If e.Button = MouseButtons.Left Then
            ReleaseCapture()
            SendMessage(Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0)
        End If
    End Sub
End Class