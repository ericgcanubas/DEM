Imports ADOX
Module ModConnection

    Public ConnServer As ADODB.Connection

    Public gbl_Database As String
    Public gbl_Server As String
    Public IsConnected As Boolean

    Public rs As ADODB.Recordset
    Public ConnLocal As New ADODB.Connection()
    Public GL_EXPORT_PATH As String
    Public Sub getConnection()
        ConnServer = New ADODB.Connection()
        Try
            With ConnServer
                .CursorLocation = ADODB.CursorLocationEnum.adUseClient
                .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" + gbl_Database + ";Data Source=" + gbl_Server
                .CommandTimeout = 60
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



    Public Function CreateData() As String

        Try
            Dim catalog As New Catalog()
            ' Create .mdb file in the specified path
            Dim strDBName As String = "Export_data"
            Dim dbPath As String = $"{GL_EXPORT_PATH}"
            Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
            catalog.Create(connectionString)
            CreateData = strDBName
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
            CreateData = ""
        End Try


    End Function

    Public Function getConString(strDBName As String) As String
        Dim dbPath As String = $"{GL_EXPORT_PATH}"
        getConString = getConnectionString(dbPath)
    End Function
    Public Function getConnectionString(dbPath As String) As String
        getConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
    End Function

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
