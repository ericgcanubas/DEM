Module ModConnection

    Public ConnMain As ADODB.Connection

    Public gbl_Database As String
    Public gbl_Server As String
    Public IsConnected As Boolean
    Public Function getConnection()
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








    End Function
    Public Function CollectionAllCashier()



    End Function


    Public Function fDateIsEmpty(sValue As String) As String


        If sValue = "" Then
            fDateIsEmpty = "null"
        Else
            fDateIsEmpty = $"'{sValue}'"
        End If
    End Function

    Public Function fSqlFormat(sValue As String) As String

        fSqlFormat = sValue.Replace("'", "`")

    End Function

End Module
