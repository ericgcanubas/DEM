﻿Imports System.IO
Imports ADOX
Module ModConnection
    Public gbl_DownloadType As Integer
    Public NItemOnly As Integer

    Public ConnServer As ADODB.Connection
    Public ConnLocal As New ADODB.Connection()

    Public ConnTemp As New ADODB.Connection()

    Public gbl_Counter As String
    Public gbl_AdjustmentOnly As Boolean
    Public gbl_Branches As Boolean
    Public gbl_Database As String
    Public gbl_Server As String
    Public IsConnected As Boolean

    Public rs As ADODB.Recordset

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

    Public Function GetParameter(ParameterName As String) As String

        Dim rx As New ADODB.Recordset
        rx.Open($"select ParameterValue from tbl_parameter WHERE ParameterName ='{ParameterName}' ", ConnTemp, ADODB.CursorTypeEnum.adOpenStatic)
        If rx.RecordCount = 0 Then
            GetParameter = ""
            SetParamter(ParameterName, "")
        Else
            GetParameter = rx.Fields("ParameterValue").Value
        End If
    End Function
    Public Sub SetParamter(ParameterName As String, ParamterValue As String)
        Dim rx As New ADODB.Recordset

        Try
            rx.Open($"select ParameterValue from tbl_parameter WHERE ParameterName ='{ParameterName}' ", ConnTemp, ADODB.CursorTypeEnum.adOpenStatic)
            If rx.RecordCount = 0 Then
                ConnTemp.Execute($"INSERT INTO tbl_parameter (ParameterName,ParameterValue) VALUES('{ParameterName}','{ParamterValue}')")
                MessageBox.Show($"NEW PARAMATER = [{ParameterName}]")
            Else
                ConnTemp.Execute($"UPDATE tbl_parameter SET ParameterValue='{ParamterValue}' WHERE ParameterName = '{ParameterName}' ")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Paramter Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
    Public Function getFilePath() As String
        Dim strDBName As String = "Main"
        getFilePath = $"{Application.StartupPath}\{strDBName}"
    End Function


    Public Function CreateDBMain()
        Dim connectionString As String = ""
        Try
            Dim catalog As New Catalog()

            Dim dbPath As String = getFilePath()
            connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
            catalog.Create(connectionString)
            CreateDBMain = connectionString

        Catch ex As Exception
            CreateDBMain = connectionString
        End Try


    End Function
    Public Function CreateDBBranch()
        Dim connectionString As String = ""
        Try
            Dim catalog As New Catalog()

            Dim strDBName As String = "Branch"

            Dim dbPath As String = $"{Application.StartupPath}\{strDBName}"
            connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
            catalog.Create(connectionString)
            CreateDBBranch = connectionString

        Catch ex As Exception
            CreateDBBranch = connectionString
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

    Public Sub Local_CreateTable_tbl_info(dt As Date, fileRef As String, Counter As String)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_info (
                                                [Counter] TEXT(5) NOT NULL,
                                                DateTransaction DATETIME NOT NULL,        
                                                DateAdded DATETIME NOT NULL,
                                                TimeAdded TEXT(30) NOT NULL,
                                                [Reference] TEXT(30) NOT NULL
                                        );"



            ConnLocal.Execute(createTableSql)

            Dim dtadded As Date = Now.Date
            Dim tmadded As TimeSpan = Now.TimeOfDay()

            ConnLocal.Execute($"INSERT INTO tbl_info 
                                            ([Counter],
                                            DateTransaction,
                                            DateAdded,
                                            TimeAdded,
                                            [Reference]) 
                                            VALUES('{Counter}',
                                            {fDateIsEmpty(dt.ToShortDateString())},
                                            '{dtadded.ToShortDateString()}',
                                            '{tmadded.ToString()}',
                                            '{fileRef}') ")


        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_info")
            Application.Exit()
        End Try
    End Sub
    Public Sub Locat_Get_info(counter As String)

        Dim rx As New ADODB.Recordset
        rx.Open($"SELECT * FROM tbl_info WHERE [COUNTER] = '{counter}'", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        If rx.RecordCount = 0 Then

        Else

        End If

    End Sub

    Public Sub SetLog(IsUpload As Boolean)
        If IsUpload = True Then

            SaveSetting("DEM", "MODE", "UPLOAD_LOG", DateTime.Now)

        Else
            SaveSetting("DEM", "MODE", "DOWNLOAD_LOG", DateTime.Now)
        End If
    End Sub
    Public Function GetLog(IsUpload As Boolean) As String
        If IsUpload = True Then
            Return GetSetting("DEM", "MODE", "UPLOAD_LOG")
        Else
            Return GetSetting("DEM", "MODE", "DOWNLOAD_LOG")
        End If
    End Function
    Public Sub FileDelete(strFile As String)

        If File.Exists(strFile) Then
            File.Delete(strFile)
        End If

    End Sub

    Public Sub FileCopy(sourcePath As String)
        Dim fileName As String = Path.GetFileName(sourcePath) ' file name
        Dim destinationPath As String = $"{Application.StartupPath}\{fileName}"

        ' Overwrite destination if it exists (optional: set to False to avoid overwrite)
        File.Copy(sourcePath, destinationPath, True)

    End Sub

End Module
