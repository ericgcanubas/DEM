Public Class FrmMainSetup
    Private Sub FrmMainSetup_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RefreshList()

        chkGenerateType.Checked = Val(GetParameter("GenerateType"))
    End Sub
    Private Sub RefreshList()
        LvLIst.Items.Clear()


        Dim rx As New ADODB.Recordset
        rx.Open("SELECT * FROM tbl_counter_list", ConnTemp, ADODB.CursorTypeEnum.adOpenStatic)



        ' Add rows to the ListView
        Do While Not rx.EOF
            ' Create a ListViewItem for the first column
            Dim item As New ListViewItem(rx.Fields(0).Value.ToString())

            ' Add the rest of the fields as subitems
            For i As Integer = 1 To rx.Fields.Count - 1
                item.SubItems.Add(rx.Fields(i).Value.ToString())
            Next

            ' Add the item to the ListView
            LvLIst.Items.Add(item)

            ' Move to next record
            rx.MoveNext()
        Loop

        ' Optional: Auto resize columns

    End Sub
    Private Sub BtnAdded_Click(sender As Object, e As EventArgs) Handles BtnAdded.Click
        Try
            If (txtCounter.Text.Length > 0) Then
                ConnTemp.Execute($"INSERT INTO tbl_counter_list ([Counter]) VALUES ('{txtCounter.Text}')")
                Application.DoEvents()
                RefreshList()
                txtCounter.Clear()
            Else

                MessageBox.Show("Counter not found")
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        RefreshList()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try

            If (LvLIst.Items.Count > 0) Then
                LvLIst.Select()

                ConnTemp.Execute($"DELETE FROM tbl_counter_list WHERE [Counter]='{LvLIst.FocusedItem.Text}'")
                Application.DoEvents()
                RefreshList()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub chkGenerateType_CheckedChanged(sender As Object, e As EventArgs) Handles chkGenerateType.CheckedChanged
        If (chkGenerateType.Checked = True) Then
            SetParamter("GenerateType", "1")
        Else
            SetParamter("GenerateType", "0")
        End If
    End Sub
End Class