Module ModConnection

    Public ConnMain As ADODB.Connection

    Public gbl_Database As String
    Public gbl_Server As String
    Public IsConnected As Boolean
    Public Sub getConnection()
        ConnMain = New ADODB.Connection()
        Try
            With ConnMain
                .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" + gbl_Database + ";Data Source=" + gbl_Server
                .Open()
                .IsolationLevel = ADODB.IsolationLevelEnum.adXactIsolated
            End With
            IsConnected = True
        Catch ex As Exception
            IsConnected = False
            MessageBox.Show(ex.Message)
            Application.Exit()

        End Try

    End Sub



    Public Function fDateIsEmpty(sValue As Object) As String
        If IsDBNull(sValue) = True Then
            fDateIsEmpty = "null"
        Else
            If sValue = "" Then
                fDateIsEmpty = "null"
            Else
                fDateIsEmpty = $"'{sValue}'"
            End If
        End If


    End Function

    Public Function fSqlFormat(sValue As Object) As String
        If IsDBNull(sValue) = True Then
            fSqlFormat = ""
        Else
            fSqlFormat = sValue.ToString().Replace("'", "`")
        End If


    End Function
    Public Function fNum(sValue As Object) As Double

        If IsDBNull(sValue) = True Then
            fNum = 0
        Else
            fNum = Val(sValue)
        End If

    End Function


End Module
