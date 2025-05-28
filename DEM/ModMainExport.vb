Imports ADOX

Module ModMainExport
    Public rs As ADODB.Recordset
    Public conn As New ADODB.Connection()
    Public GL_EXPORT_PATH As String
    Public Function CreateSmallDatabase() As String

        Try
            Dim catalog As New Catalog()
            ' Create .mdb file in the specified path
            Dim strDBName As String = "Export_data"
            Dim dbPath As String = $"{GL_EXPORT_PATH}"
            Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
            catalog.Create(connectionString)
            CreateSmallDatabase = strDBName
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
            CreateSmallDatabase = ""
        End Try


    End Function

    Public Function getConString(strDBName As String) As String

        Dim dbPath As String = $"{GL_EXPORT_PATH}"
        getConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath

    End Function
    Public Sub CreateTable_tbl_PCPOS_Cashiers()

        Try
            Dim createTableSql As String = "
            CREATE TABLE tbl_PCPOS_Cashiers (
                CashierCode TEXT(3) NOT NULL,
                [Password] TEXT(3) NOT NULL,
                Senior BYTE NOT NULL,
                Track2 TEXT(60),
                Track1 TEXT(60),
                DirectVoid BYTE NOT NULL,
                DirectDiscount BYTE NOT NULL,
                DirectSurcharge BYTE NOT NULL,
                SecureCode TEXT(10),
                FullName TEXT(30),
                CodeType INTEGER,
                DiscountLimit DOUBLE NOT NULL,
                Active BYTE NOT NULL,
                [PK] INTEGER,
                Changes BYTE NOT NULL,
                Admin BYTE NOT NULL,
                Transfered BYTE NOT NULL
            )"


            conn.Execute(createTableSql)

        Catch ex As Exception
            MessageBox.Show(ex.Message, " tbl_PCPOS_Cashiers")
            Application.Exit()
        End Try







    End Sub

    Public Sub Collect_tbl_PCPOS_Cashiers(pb As ProgressBar, l As Label)

        Try
            rs = New ADODB.Recordset
            rs.Open("select * from tbl_PCPOS_Cashiers ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
            pb.Maximum = rs.RecordCount
            pb.Value = 0
            pb.Minimum = 0
            If rs.RecordCount > 0 Then

                While Not rs.EOF
                    pb.Value = pb.Value + 1
                    l.Text = pb.Maximum & "/" & pb.Value
                    Application.DoEvents()
                    conn.Execute($"INSERT INTO tbl_PCPOS_Cashiers 
                                    (CashierCode,
                                    [Password],
                                    Senior,
                                    Track2,
                                    Track1,
                                    DirectVoid,
                                    DirectDiscount,
                                    DirectSurcharge,
                                    SecureCode,
                                    FullName,
                                    CodeType,
                                    DiscountLimit,
                                    Active,
                                    [PK],
                                    Changes,
                                    Admin,
                                    Transfered )   
                            VALUES ('{rs.Fields("CashierCode").Value}',
                                    '{rs.Fields("Password").Value}',
                                    {rs.Fields("Senior").Value},
                                    '{rs.Fields("Track2").Value}',
                                    '{rs.Fields("Track1").Value}',
                                    {rs.Fields("DirectVoid").Value},
                                    {rs.Fields("DirectDiscount").Value},
                                    {rs.Fields("DirectSurcharge").Value},
                                    '{rs.Fields("SecureCode").Value}',
                                    '{rs.Fields("FullName").Value}',
                                    {rs.Fields("CodeType").Value},
                                    {rs.Fields("DiscountLimit").Value},
                                    {rs.Fields("Active").Value},
                                    {rs.Fields("PK").Value},
                                    {rs.Fields("Changes").Value},
                                    {rs.Fields("Admin").Value},
                                    {rs.Fields("Transfered").Value}
                                )        

            ")
                    rs.MoveNext()
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, " tbl_PCPOS_Cashiers")
            Application.Exit()
        End Try



    End Sub

    Public Sub CreateTable_tbl_ItemsForPLU()
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_ItemsForPLU (
                                            ItemCode TEXT(12),
                                            ECRDescription TEXT(45),
                                            ItemDescription TEXT(45),
                                            GrossSRP DOUBLE,
                                            PromoDisc DOUBLE,
                                            PromoFrom DATETIME,
                                            PromoTo DATETIME
                                )"

            conn.Execute(createTableSql)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_ItemsForPLU")
            Application.Exit()
        End Try

    End Sub
    Public Sub Collect_tbl_ItemsForPLU(pb As ProgressBar, l As Label)


        Try
            rs = New ADODB.Recordset
            rs.Open("select tbl_ItemsForPLU.*  FROM tbl_ItemsForPLU inner join  tbl_Items on  [tbl_Items].ItemCode = tbl_ItemsForPLU.ItemCode where [tbl_Items].[status] = 0 ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
            pb.Maximum = rs.RecordCount
            pb.Value = 0
            pb.Minimum = 0
            If rs.RecordCount > 0 Then
                While Not rs.EOF
                    pb.Value = pb.Value + 1
                    l.Text = pb.Maximum & "/" & pb.Value
                    Application.DoEvents()
                    Dim strSQL As String = $" INSERT INTO tbl_ItemsForPLU 
                                    (ItemCode,
                                    ECRDescription,
                                    ItemDescription,
                                    GrossSRP,
                                    PromoDisc,
                                    PromoFrom,
                                    PromoTo )  
                                    VALUES ('{rs.Fields("ItemCode").Value}',
                                    '{fSqlFormat(rs.Fields("ECRDescription").Value)}',
                                    '{fSqlFormat(rs.Fields("ItemDescription").Value)}',
                                     {rs.Fields("GrossSRP").Value},
                                     {rs.Fields("PromoDisc").Value},
                                     {fDateIsEmpty(rs.Fields("PromoFrom").Value.ToString())},
                                     {fDateIsEmpty(rs.Fields("PromoTo").Value.ToString())}
                                );"
                    conn.Execute(strSQL)
                    rs.MoveNext()
                End While

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_ItemsForPLU")
            Application.Exit()
        End Try

    End Sub

    Public Sub CreateTable_tbl_bank()



        Try
            Dim createTableSql As String = " CREATE Table tbl_Bank (
                                                PK INTEGER PRIMARY KEY,
                                                BankName Text(50) Not NULL,
                                                [Address] Text(50) Not NULL,
                                                TelNo Text(50) Not NULL,
                                                FaxNo Text(50) Not NULL,
                                                ContactPerson Text(100) Not NULL,
                                                LastModified Text(50),
                                                Tax Double Not NULL,
                                                Locked Byte Not NULL,
                                                CardType Integer Not NULL,
                                                IsDefault Integer Not NULL
                                );"

            conn.Execute(createTableSql)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Bank")
            Application.Exit()
        End Try
    End Sub
    Public Sub Collect_tbl_Bank(pb As ProgressBar, l As Label)


        Try
            rs = New ADODB.Recordset
            rs.Open("select * from tbl_Bank ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
            pb.Maximum = rs.RecordCount
            pb.Value = 0
            pb.Minimum = 0
            If rs.RecordCount > 0 Then
                While Not rs.EOF
                    pb.Value = pb.Value + 1
                    l.Text = pb.Maximum & "/" & pb.Value
                    Application.DoEvents()
                    Dim strSQL As String = $" INSERT INTO tbl_Bank 
                                    (PK,
                                    BankName,
                                    [Address],
                                    TelNo,
                                    FaxNo,
                                    ContactPerson,
                                    LastModified,
                                    Tax,
                                    Locked,
                                    CardType,
                                    IsDefault)  
                                    VALUES ({rs.Fields("PK").Value},
                                    '{fSqlFormat(rs.Fields("BankName").Value)}',
                                    '{fSqlFormat(rs.Fields("Address").Value)}',
                                    '{fSqlFormat(rs.Fields("TelNo").Value)}',
                                    '{fSqlFormat(rs.Fields("FaxNo").Value)}',
                                    '{fSqlFormat(rs.Fields("ContactPerson").Value)}',
                                    '{fSqlFormat(rs.Fields("LastModified").Value)}',
                                    {rs.Fields("Tax").Value},
                                    {rs.Fields("Locked").Value},
                                    {rs.Fields("CardType").Value},
                                    {rs.Fields("IsDefault").Value}
                            
                                );"
                    conn.Execute(strSQL)
                    rs.MoveNext()
                End While

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Bank")
            Application.Exit()
        End Try

    End Sub

    Public Sub CreateTable_tbl_banks()
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_Banks (
                                            PK INTEGER PRIMARY KEY,
                                            BankCode TEXT(2) NOT NULL,
                                            BankName TEXT(50) NOT NULL,
                                            Telephone TEXT(50) NOT NULL,
                                            MERC_COD TEXT(50) NOT NULL,
                                            MERC_COD2 TEXT(50) NOT NULL,
                                            [Description] TEXT(12) NOT NULL,
                                            Bank INTEGER NOT NULL
                                        );"
            conn.Execute(createTableSql)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Banks")
            Application.Exit()
        End Try
    End Sub
    Public Sub Collect_tbl_Banks(pb As ProgressBar, l As Label)
        Try
            rs = New ADODB.Recordset
            rs.Open("select * from tbl_Banks ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
            pb.Maximum = rs.RecordCount
            pb.Value = 0
            pb.Minimum = 0
            If rs.RecordCount > 0 Then
                While Not rs.EOF
                    pb.Value = pb.Value + 1
                    l.Text = pb.Maximum & "/" & pb.Value
                    Application.DoEvents()
                    Dim strSQL As String = $"INSERT INTO tbl_Banks 
                                    (PK,
                                    BankCode,
                                    BankName,
                                    Telephone,
                                    MERC_COD,
                                    MERC_COD2,
                                    [Description],
                                    Bank)
                                    VALUES ({rs.Fields("PK").Value},
                                    '{fSqlFormat(rs.Fields("BankCode").Value)}',
                                    '{fSqlFormat(rs.Fields("BankName").Value)}',
                                    '{fSqlFormat(rs.Fields("Telephone").Value)}',
                                    '{fSqlFormat(rs.Fields("MERC_COD").Value)}',
                                    '{fSqlFormat(rs.Fields("MERC_COD2").Value)}',
                                    '{fSqlFormat(rs.Fields("Description").Value)}',
                                     {rs.Fields("Bank").Value}       
                                );"
                    conn.Execute(strSQL)
                    rs.MoveNext()
                End While

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Banks")
            Application.Exit()
        End Try

    End Sub
    Public Sub CreateTable_tbl_Bank_Terms()
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_Bank_Terms (
                                            BankKey INTEGER NOT NULL,
                                            Effectivity DATETIME NOT NULL,
                                            [Type] TEXT(50) NOT NULL,
                                            Terms TEXT(50) NOT NULL,
                                            TermsDescription TEXT(255) NOT NULL
                                        );"
            conn.Execute(createTableSql)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Bank_Terms")
            Application.Exit()
        End Try
    End Sub

    Public Sub Collect_tbl_Bank_Terms(pb As ProgressBar, l As Label)

        Try
            rs = New ADODB.Recordset
            rs.Open("select * from tbl_Bank_Terms ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
            pb.Maximum = rs.RecordCount
            pb.Value = 0
            pb.Minimum = 0
            If rs.RecordCount > 0 Then
                While Not rs.EOF
                    pb.Value = pb.Value + 1
                    l.Text = pb.Maximum & "/" & pb.Value
                    Application.DoEvents()
                    Dim strSQL As String = $"INSERT INTO tbl_Bank_Terms 
                                    (BankKey,
                                    Effectivity,
                                    Type,
                                    Terms,
                                    TermsDescription)
                                    VALUES ({rs.Fields("BankKey").Value},
                                    {fDateIsEmpty(rs.Fields("Effectivity").Value.ToString())},
                                    '{fSqlFormat(rs.Fields("Type").Value)}',
                                    '{fSqlFormat(rs.Fields("Terms").Value)}',
                                    '{fSqlFormat(rs.Fields("TermsDescription").Value)}'
                                );"
                    conn.Execute(strSQL)
                    rs.MoveNext()
                End While

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Bank_Terms")
            Application.Exit()
        End Try

    End Sub
    Public Sub CreateTable_tbl_QRPay_Type()
        Try
            Dim createTableSql As String = " CREATE TABLE tbl_QRPay_Type (
                                                nQRPTypeID INTEGER PRIMARY KEY,
                                                sQRType TEXT(50),
                                                nPercRate Double,
                                                nSort INTEGER
                                            );"

            conn.Execute(createTableSql)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_QRPay_Type")
            Application.Exit()
        End Try
    End Sub
    Public Sub Collect_tbl_QRPay_Type(pb As ProgressBar, l As Label)

        Try
            rs = New ADODB.Recordset
            rs.Open("select * from tbl_QRPay_Type ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
            pb.Maximum = rs.RecordCount
            pb.Value = 0
            pb.Minimum = 0
            If rs.RecordCount > 0 Then
                While Not rs.EOF
                    pb.Value = pb.Value + 1
                    l.Text = pb.Maximum & "/" & pb.Value
                    Application.DoEvents()
                    Dim strSQL As String = $"INSERT INTO tbl_QRPay_Type 
                                    (nQRPTypeID,
                                    sQRType,
                                    nPercRate,
                                    nSort)
                                    VALUES ({rs.Fields("nQRPTypeID").Value},
                                    '{fSqlFormat(rs.Fields("sQRType").Value.ToString())}',
                                    {rs.Fields("nPercRate").Value},
                                    {rs.Fields("nSort").Value}
                                );"
                    conn.Execute(strSQL)
                    rs.MoveNext()
                End While

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_QRPay_Type")
            Application.Exit()
        End Try

    End Sub

    Public Sub CreateTable_tbl_GiftCert_List()
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_GiftCert_List (
                                                PK INTEGER PRIMARY KEY,
                                                GCNumber DOUBLE NOT NULL,
                                                Amount DOUBLE NOT NULL,
                                                Customer TEXT(255) NOT NULL,
                                                ValidFrom DATETIME NOT NULL,
                                                ValidTo DATETIME NOT NULL,
                                                DateAdded DATETIME NOT NULL,
                                                Used BYTE NOT NULL,
                                                DateUsed DATETIME
                                        );"

            conn.Execute(createTableSql)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_GiftCert_List")
            Application.Exit()
        End Try
    End Sub

    Public Sub Collect_tbl_GiftCert_List(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 1
        Try
            rs = New ADODB.Recordset
            rs.Open($"select * from tbl_GiftCert_List where YEAR(ValidTo) > {year}  and DateUsed is null ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
            pb.Maximum = rs.RecordCount
            pb.Value = 0
            pb.Minimum = 0
            If rs.RecordCount > 0 Then
                While Not rs.EOF
                    pb.Value = pb.Value + 1
                    l.Text = pb.Maximum & "/" & pb.Value
                    Application.DoEvents()
                    Dim strSQL As String = $"INSERT INTO tbl_GiftCert_List 
                                    (PK,
                                    GCNumber,
                                    Amount,
                                    Customer,
                                    ValidFrom,
                                    ValidTo,
                                    DateAdded,
                                    Used,
                                    DateUsed)
                                    VALUES ({rs.Fields("PK").Value},      
                                    {rs.Fields("GCNumber").Value},
                                    {rs.Fields("Amount").Value},
                                   '{fSqlFormat(rs.Fields("Customer").Value.ToString())}',
                                    {fDateIsEmpty(rs.Fields("ValidFrom").Value.ToString())},
                                    {fDateIsEmpty(rs.Fields("ValidTo").Value.ToString())},
                                    {fDateIsEmpty(rs.Fields("DateAdded").Value.ToString())},
                                    {rs.Fields("Used").Value},
                                    {fDateIsEmpty(rs.Fields("DateUsed").Value.ToString())}
                                );"
                    conn.Execute(strSQL)
                    rs.MoveNext()
                End While

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_GiftCert_List")
            Application.Exit()
        End Try
    End Sub
    Public Sub CreateTable_tbl_VPlus_Codes()
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_VPlus_Codes (
                                            Codes TEXT(16) NOT NULL,
                                            Customer INTEGER,
                                            InPoints DOUBLE NOT NULL,
                                            OutPoints DOUBLE NOT NULL,
                                            AvailPoints DOUBLE,
                                            Blocked BYTE NOT NULL,
                                            Printed BYTE NOT NULL,
                                            CreatedOn DATETIME NOT NULL,
                                            CreatedOnTime DATETIME NOT NULL,
                                            [Password] TEXT(6) NOT NULL,
                                            DateStarted DATETIME,
                                            DateExpired DATETIME,
                                            DateModified DATETIME NOT NULL,
                                            Changes BYTE NOT NULL
                                        );"

            conn.Execute(createTableSql)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_VPlus_Codes")
            Application.Exit()
        End Try
    End Sub
    Public Sub Collect_tbl_VPlus_Codes(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 5
        Try
            rs = New ADODB.Recordset
            rs.Open($"select * from tbl_VPlus_Codes where year(DateExpired) > {year} ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
            pb.Maximum = rs.RecordCount
            pb.Value = 0
            pb.Minimum = 0
            If rs.RecordCount > 0 Then
                While Not rs.EOF
                    pb.Value = pb.Value + 1
                    l.Text = pb.Maximum & "/" & pb.Value
                    Application.DoEvents()
                    Dim strSQL As String = $"INSERT INTO tbl_VPlus_Codes 
                                    (Codes,
                                    Customer,
                                    InPoints,
                                    OutPoints,
                                    AvailPoints,
                                    Blocked,
                                    Printed,
                                    CreatedOn,
                                    CreatedOnTime,
                                    [Password],
                                    DateStarted,
                                    DateExpired,
                                    DateModified,
                                    Changes)
                                    VALUES ('{fSqlFormat(rs.Fields("Codes").Value)}',      
                                    {fNum(rs.Fields("Customer").Value)},
                                    {fNum(rs.Fields("InPoints").Value)},
                                    {fNum(rs.Fields("OutPoints").Value)},
                                    {fNum(rs.Fields("AvailPoints").Value)},
                                    {rs.Fields("Blocked").Value},
                                    {rs.Fields("Printed").Value},
                                    {fDateIsEmpty(rs.Fields("CreatedOn").Value.ToString())},
                                    {fDateIsEmpty(rs.Fields("CreatedOnTime").Value.ToString())},
                                   '{fSqlFormat(rs.Fields("Password").Value.ToString())}',
                                    {fDateIsEmpty(rs.Fields("DateStarted").Value.ToString())},
                                    {fDateIsEmpty(rs.Fields("DateExpired").Value.ToString())},
                                    {fDateIsEmpty(rs.Fields("DateModified").Value.ToString())},
                                    {rs.Fields("Changes").Value} );"
                    conn.Execute(strSQL)
                    rs.MoveNext()
                End While

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_VPlus_Codes")
            Application.Exit()
        End Try
    End Sub

    Public Sub CreateTable_tbl_VPlus_Codes_Validity()
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_VPlus_Codes_Validity (
                                            Codes TEXT(16) NOT NULL,
                                            DateStarted DATETIME NOT NULL,
                                            DateExpired DATETIME NOT NULL,
                                            GracePeriod DATETIME NOT NULL
                                        );"

            conn.Execute(createTableSql)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_VPlus_Codes_Validity")
            Application.Exit()
        End Try
    End Sub
    Public Sub Collect_tbl_VPlus_Codes_Validity(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 5
        Try
            rs = New ADODB.Recordset
            rs.Open($"select tbl_VPlus_Codes_Validity.* from tbl_VPlus_Codes_Validity join  tbl_VPlus_Codes on tbl_VPlus_Codes.codes = tbl_VPlus_Codes_Validity.codes  where year(tbl_VPlus_Codes.DateExpired) > {year} ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
            pb.Maximum = rs.RecordCount
            pb.Value = 0
            pb.Minimum = 0
            If rs.RecordCount > 0 Then
                While Not rs.EOF
                    pb.Value = pb.Value + 1
                    l.Text = pb.Maximum & "/" & pb.Value
                    Application.DoEvents()
                    Dim strSQL As String = $"INSERT INTO tbl_VPlus_Codes_Validity 
                                    (Codes,
                                    DateStarted,
                                    DateExpired,
                                    GracePeriod)
                                    VALUES ('{fSqlFormat(rs.Fields("Codes").Value)}',  
                                    {fDateIsEmpty(rs.Fields("DateStarted").Value.ToString())},
                                    {fDateIsEmpty(rs.Fields("DateExpired").Value.ToString())},    
                                    {fDateIsEmpty(rs.Fields("GracePeriod").Value.ToString())}   
                                   );"

                    conn.Execute(strSQL)
                    rs.MoveNext()
                End While

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_VPlus_Codes_Validity")
            Application.Exit()
        End Try
    End Sub


    Public Sub CreateTable_tbl_Bank_Changes()
        Try
            Dim createTableSql As String = "  CREATE TABLE tbl_Bank_Changes (
                                                PK INTEGER PRIMARY KEY,
                                                EffectDate DATETIME,
                                                BankKey Integer,
                                                [Changes] TEXT(50)
                                            );"

            conn.Execute(createTableSql)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Bank_Changes")
            Application.Exit()
        End Try
    End Sub
    Public Sub Collect_tbl_Bank_Changes(pb As ProgressBar, l As Label)

        Try
            rs = New ADODB.Recordset
            rs.Open($"select * from tbl_Bank_Changes", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
            pb.Maximum = rs.RecordCount
            pb.Value = 0
            pb.Minimum = 0
            If rs.RecordCount > 0 Then
                While Not rs.EOF
                    pb.Value = pb.Value + 1
                    l.Text = pb.Maximum & "/" & pb.Value
                    Application.DoEvents()
                    Dim strSQL As String = $"INSERT INTO tbl_Bank_Changes 
                                    (PK,
                                    EffectDate,
                                    BankKey,
                                    [Changes])
                                    VALUES ('{fSqlFormat(rs.Fields("PK").Value)}',  
                                    {fDateIsEmpty(rs.Fields("EffectDate").Value.ToString())},
                                    {fNum(rs.Fields("BankKey").Value)},    
                                    '{fSqlFormat(rs.Fields("Changes").Value)}'   
                                   );"

                    conn.Execute(strSQL)
                    rs.MoveNext()
                End While

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Bank_Changes")
            Application.Exit()
        End Try
    End Sub

    Public Sub CreateTable_tbl_PCPOS_Cashiers_Changes()
        Try
            Dim createTableSql As String = " CREATE TABLE tbl_PCPOS_Cashiers_Changes (
                                            PK INTEGER PRIMARY KEY,
                                            EffectDate DATETIME,
                                            CashierPK INTEGER,
                                            [Changes] TEXT(50)
                                        );"

            conn.Execute(createTableSql)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_PCPOS_Cashiers_Changes")
            Application.Exit()
        End Try
    End Sub
    Public Sub Collect_tbl_PCPOS_Cashiers_Changes(pb As ProgressBar, l As Label)

        Try
            rs = New ADODB.Recordset
            rs.Open($"select * from tbl_PCPOS_Cashiers_Changes", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
            pb.Maximum = rs.RecordCount
            pb.Value = 0
            pb.Minimum = 0
            If rs.RecordCount > 0 Then
                While Not rs.EOF
                    pb.Value = pb.Value + 1
                    l.Text = pb.Maximum & "/" & pb.Value
                    Application.DoEvents()
                    Dim strSQL As String = $"INSERT INTO tbl_PCPOS_Cashiers_Changes 
                                    (PK,
                                    EffectDate,
                                    CashierPK,
                                    [Changes])
                                    VALUES ('{fSqlFormat(rs.Fields("PK").Value)}',  
                                    {fDateIsEmpty(rs.Fields("EffectDate").Value.ToString())},
                                    {fNum(rs.Fields("CashierPK").Value)},    
                                    '{fSqlFormat(rs.Fields("Changes").Value)}'   
                                   );"

                    conn.Execute(strSQL)
                    rs.MoveNext()
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_PCPOS_Cashiers_Changes")
            Application.Exit()
        End Try
    End Sub
End Module
