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
    Public Sub CreateTable_tbl_PCPOS_Cashiers(pb As ProgressBar, l As Label)

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
            Collect_tbl_PCPOS_Cashiers(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, " tbl_PCPOS_Cashiers")
            Application.Exit()
        End Try
    End Sub

    Private Sub Collect_tbl_PCPOS_Cashiers(pb As ProgressBar, l As Label)

        rs = New ADODB.Recordset
        rs.Open("select * from tbl_PCPOS_Cashiers ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then

            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PCPOS_Cashiers :" & pb.Maximum & "/" & pb.Value
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



    End Sub

    Public Sub CreateTable_tbl_ItemsForPLU(pb As ProgressBar, l As Label)
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
            Collect_tbl_ItemsForPLU(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_ItemsForPLU")
            Application.Exit()
        End Try

    End Sub
    Private Sub Collect_tbl_ItemsForPLU(pb As ProgressBar, l As Label)
        rs = New ADODB.Recordset
        rs.Open("select tbl_ItemsForPLU.*  FROM tbl_ItemsForPLU inner join  tbl_Items on  [tbl_Items].ItemCode = tbl_ItemsForPLU.ItemCode join tbl_Suppliers on tbl_Suppliers.PK = tbl_Items.SupplierKey  where [tbl_Items].[status] = 0 and tbl_Suppliers.SStatus = 0 ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        Dim n As Integer = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_ItemsForPLU :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
                n = n + 1
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


    End Sub

    Public Sub CreateTable_tbl_bank(pb As ProgressBar, l As Label)

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
            Collect_tbl_Bank(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Bank")
            Application.Exit()
        End Try
    End Sub
    Private Sub Collect_tbl_Bank(pb As ProgressBar, l As Label)



        rs = New ADODB.Recordset
        rs.Open("select * from tbl_Bank ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Bank :" & pb.Maximum & "/" & pb.Value
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


    End Sub

    Public Sub CreateTable_tbl_banks(pb As ProgressBar, l As Label)
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
            Collect_tbl_Banks(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Banks")
            Application.Exit()
        End Try
    End Sub
    Private Sub Collect_tbl_Banks(pb As ProgressBar, l As Label)

        rs = New ADODB.Recordset
        rs.Open("select * from tbl_Banks ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Banks :" & pb.Maximum & "/" & pb.Value
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


    End Sub

    Public Sub CreateTable_tbl_Banks_Changes(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = "  CREATE TABLE tbl_Banks_Changes (
                                                PK INTEGER PRIMARY KEY,
                                                EffectDate DATETIME,
                                                BankKey Integer,
                                                [Changes] TEXT(50)
                                            );"

            conn.Execute(createTableSql)
            Collect_tbl_Banks_Changes(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Banks_Changes")
            Application.Exit()
        End Try
    End Sub
    Private Sub Collect_tbl_Banks_Changes(pb As ProgressBar, l As Label)


        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_Banks_Changes ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Banks_Changes :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim strSQL As String = $"INSERT INTO tbl_Banks_Changes 
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

    End Sub



    Public Sub CreateTable_tbl_Bank_Changes(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = "  CREATE TABLE tbl_Bank_Changes (
                                                PK INTEGER PRIMARY KEY,
                                                EffectDate DATETIME,
                                                BankKey Integer,
                                                [Changes] TEXT(50)
                                            );"

            conn.Execute(createTableSql)
            Collect_tbl_Bank_Changes(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Bank_Changes")
            Application.Exit()
        End Try
    End Sub
    Private Sub Collect_tbl_Bank_Changes(pb As ProgressBar, l As Label)


        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_Bank_Changes", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Bank_Changes :" & pb.Maximum & "/" & pb.Value
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

    End Sub

    Public Sub CreateTable_tbl_Bank_Terms(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_Bank_Terms (
                                            BankKey INTEGER NOT NULL,
                                            Effectivity DATETIME NOT NULL,
                                            [Type] TEXT(50) NOT NULL,
                                            Terms TEXT(50) NOT NULL,
                                            TermsDescription TEXT(255) NOT NULL
                                        );"
            conn.Execute(createTableSql)
            Collect_tbl_Bank_Terms(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Bank_Terms")
            Application.Exit()
        End Try
    End Sub

    Private Sub Collect_tbl_Bank_Terms(pb As ProgressBar, l As Label)


        rs = New ADODB.Recordset
        rs.Open("select * from tbl_Bank_Terms ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Bank_Terms :" & pb.Maximum & "/" & pb.Value
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


    End Sub
    Public Sub CreateTable_tbl_QRPay_Type(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = " CREATE TABLE tbl_QRPay_Type (
                                                nQRPTypeID INTEGER PRIMARY KEY,
                                                sQRType TEXT(50),
                                                nPercRate Double,
                                                nSort INTEGER
                                            );"

            conn.Execute(createTableSql)
            Collect_tbl_QRPay_Type(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_QRPay_Type")
            Application.Exit()
        End Try
    End Sub
    Private Sub Collect_tbl_QRPay_Type(pb As ProgressBar, l As Label)


        rs = New ADODB.Recordset
        rs.Open("select * from tbl_QRPay_Type ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_QRPay_Type :" & pb.Maximum & "/" & pb.Value
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


    End Sub

    Public Sub CreateTable_tbl_GiftCert_List(pb As ProgressBar, l As Label)
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
            Collect_tbl_GiftCert_List(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_GiftCert_List")
            Application.Exit()
        End Try
    End Sub

    Private Sub Collect_tbl_GiftCert_List(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 1

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_GiftCert_List where YEAR(ValidTo) > {year}  and DateUsed is null ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_GiftCert_List :" & pb.Maximum & "/" & pb.Value
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

    End Sub

    Public Sub CreateTable_tbl_VPlus_Codes(pb As ProgressBar, l As Label)
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
            Collect_tbl_VPlus_Codes(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_VPlus_Codes")
            Application.Exit()
        End Try
    End Sub
    Private Sub Collect_tbl_VPlus_Codes(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 5

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_VPlus_Codes where year(DateExpired) > {year} ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_VPlus_Codes :" & pb.Maximum & "/" & pb.Value
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

    End Sub

    Public Sub CreateTable_tbl_VPlus_Codes_Validity(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_VPlus_Codes_Validity (
                                            Codes TEXT(16) NOT NULL,
                                            DateStarted DATETIME NOT NULL,
                                            DateExpired DATETIME NOT NULL,
                                            GracePeriod DATETIME NOT NULL
                                        );"

            conn.Execute(createTableSql)
            Collect_tbl_VPlus_Codes_Validity(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_VPlus_Codes_Validity")
            Application.Exit()
        End Try
    End Sub
    Private Sub Collect_tbl_VPlus_Codes_Validity(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 5

        rs = New ADODB.Recordset
        rs.Open($"select tbl_VPlus_Codes_Validity.* from tbl_VPlus_Codes_Validity join tbl_VPlus_Codes on tbl_VPlus_Codes.codes = tbl_VPlus_Codes_Validity.codes  where year(tbl_VPlus_Codes.DateExpired) > {year} ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_VPlus_Codes_Validity :" & pb.Maximum & "/" & pb.Value
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

    End Sub


    Public Sub CreateTable_tbl_PCPOS_Cashiers_Changes(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = " CREATE TABLE tbl_PCPOS_Cashiers_Changes (
                                            PK INTEGER PRIMARY KEY,
                                            EffectDate DATETIME,
                                            CashierPK INTEGER,
                                            [Changes] TEXT(50)
                                        );"

            conn.Execute(createTableSql)
            Collect_tbl_PCPOS_Cashiers_Changes(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_PCPOS_Cashiers_Changes")
            Application.Exit()
        End Try
    End Sub
    Private Sub Collect_tbl_PCPOS_Cashiers_Changes(pb As ProgressBar, l As Label)
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PCPOS_Cashiers_Changes", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PCPOS_Cashiers_Changes :" & pb.Maximum & "/" & pb.Value
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

    End Sub

    Public Sub CreateTable_tbl_Items_Change(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = " CREATE TABLE tbl_Items_Change (
                                                PK INTEGER PRIMARY KEY,
                                                ItemCode TEXT(12),
                                                ItemDescription TEXT(45),
                                                GrossSRP DOUBLE,
                                                DateChange DATETIME,
                                                Remarks TEXT(50),
                                                UserName TEXT(50),
                                                DateTimeChange DATETIME,
                                                ItemKey INTEGER
                                            );"

            conn.Execute(createTableSql)
            Collect_tbl_Items_Changes(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Items_Change")
            Application.Exit()
        End Try
    End Sub

    Private Sub Collect_tbl_Items_Changes(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 1
        rs = New ADODB.Recordset
        rs.Open($"select top 10000 * from tbl_Items_Change where year(DateChange) >= {year} order by dateChange desc", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Items_Change :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim strSQL As String = $"INSERT INTO tbl_Items_Change 
                                    (PK,
                                    ItemCode,
                                    ItemDescription,
                                    GrossSRP,
                                    DateChange,
                                    Remarks,
                                    UserName,
                                    DateTimeChange,
                                    ItemKey)
                                    VALUES ({rs.Fields("PK").Value},  
                                    '{fSqlFormat(rs.Fields("ItemCode").Value)}',
                                    '{fSqlFormat(rs.Fields("ItemDescription").Value)}',     
                                     {fNum(rs.Fields("GrossSRP").Value)},   
                                     {fDateIsEmpty(rs.Fields("DateChange").Value.ToString())},
                                     '{fSqlFormat(rs.Fields("Remarks").Value)}',
                                     '{fSqlFormat(rs.Fields("UserName").Value)}',
                                      {fDateIsEmpty(rs.Fields("DateTimeChange").Value.ToString())},
                                      {fNum(rs.Fields("ItemKey").Value)}                         
                                   );"

                conn.Execute(strSQL)
                rs.MoveNext()
            End While
        End If
    End Sub
    Public Sub CreateTable_tbl_ItemsForPLU_For_Effect(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = " CREATE TABLE tbl_ItemsForPLU_For_Effect (
                                                    PK INTEGER PRIMARY KEY,
                                                    EffectDate DATETIME,
                                                    ItemCode TEXT(12),
                                                    ItemDescription TEXT(45),
                                                    GrossSRP DOUBLE,
                                                    PromoDisc DOUBLE,
                                                    PromoFrom DATETIME,
                                                    PromoTo DATETIME
                                                );"

            conn.Execute(createTableSql)
            Collect_tbl_ItemsForPLU_For_Effect(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_ItemsForPLU_For_Effect")
            Application.Exit()
        End Try
    End Sub

    Private Sub Collect_tbl_ItemsForPLU_For_Effect(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 1
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_ItemsForPLU_For_Effect", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0

        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_ItemsForPLU_For_Effect :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim strSQL As String = $"INSERT INTO tbl_ItemsForPLU_For_Effect 
                                    (PK,
                                    EffectDate,
                                    ItemCode,
                                    ItemDescription,
                                    GrossSRP,
                                    PromoDisc,
                                    PromoFrom,
                                    PromoTo)
                                    VALUES ({rs.Fields("PK").Value},  
                                    {fDateIsEmpty(rs.Fields("EffectDate").Value.ToString())},
                                    '{fSqlFormat(rs.Fields("ItemCode").Value)}',
                                    '{fSqlFormat(rs.Fields("ItemDescription").Value)}',     
                                    {fNum(rs.Fields("GrossSRP").Value)},   
                                    {fNum(rs.Fields("PromoDisc").Value)},
                                    {fDateIsEmpty(rs.Fields("PromoFrom").Value.ToString())},       
                                    {fDateIsEmpty(rs.Fields("PromoTo").Value.ToString())}                    
                                   );"

                conn.Execute(strSQL)
                rs.MoveNext()


            End While
        End If
    End Sub

    Public Sub CreateTable_tbl_Items(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_Items (
                                            PK INTEGER PRIMARY KEY,
                                            ItemCode TEXT(12) NOT NULL,
                                            ItemDescription TEXT(45) NOT NULL,
                                            ItemType BYTE NOT NULL,
                                            ECRDescription TEXT(16),
                                            StockNumber TEXT(15),
                                            UnitOfMeasure TEXT(3),
                                            ClassKey INTEGER NOT NULL,
                                            SupplierKey INTEGER NOT NULL,
                                            Discount TEXT(15),
                                            Commission TEXT(15),
                                            Terms TEXT(15),
                                            Remarks TEXT(15),
                                            ForeignCost TEXT(50),
                                            GrossCost CURRENCY NOT NULL,
                                            Vat DOUBLE NOT NULL,
                                            MarkUp DOUBLE NOT NULL,
                                            GrossSRP CURRENCY NOT NULL,
                                            LastModifiedBy TEXT(50),
                                            PhasedOut BYTE NOT NULL,
                                            BrandKey INTEGER NOT NULL,
                                            ProdLineKey INTEGER NOT NULL,
                                            OldCode TEXT(50),
                                            SeasonCode TEXT(50),
                                            [Change] BYTE,
                                            MinQty DOUBLE NOT NULL,
                                            MaxQty DOUBLE NOT NULL,
                                            ReOrder BYTE NOT NULL,
                                            Category INTEGER NOT NULL,
                                            PromoDisc DOUBLE NOT NULL,
                                            PromoDiscAmt DOUBLE NOT NULL,
                                            PromoFrom DATETIME,
                                            PromoTo DATETIME,
                                            PromoDiscLocked BYTE NOT NULL,
                                            Level1 DOUBLE NOT NULL,
                                            Level2 DOUBLE NOT NULL,
                                            Level3 DOUBLE NOT NULL,
                                            Level4 DOUBLE NOT NULL,
                                            Level5 DOUBLE NOT NULL,
                                            Disc1 DOUBLE NOT NULL,
                                            Disc2 DOUBLE NOT NULL,
                                            Disc3 DOUBLE NOT NULL,
                                            Disc4 DOUBLE NOT NULL,
                                            Disc5 DOUBLE NOT NULL,
                                            LastCost DOUBLE,
                                            LastSRP DOUBLE,
                                            Color TEXT(255),
                                            StoreLocation INTEGER NOT NULL,
                                            PO INTEGER NOT NULL,
                                            Date_Encoded DATETIME NOT NULL,
                                            User_Action TEXT(50),
                                            User_Encoded TEXT(50),
                                            Changes TEXT(255),
                                            RefNoID INTEGER,
                                            NotIncludeInSale BYTE NOT NULL,
                                            Active BYTE,
                                            ActiveAsOf DATETIME,
                                            Discounted BYTE,
                                            MarkDown BYTE,
                                            Status BYTE NOT NULL
                                        );"

            conn.Execute(createTableSql)
            Collect_tbl_Items(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Items")
            Application.Exit()
        End Try
    End Sub

    Private Sub Collect_tbl_Items(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 1

        rs = New ADODB.Recordset
        rs.Open($"select i.* from tbl_Items as i join tbl_Suppliers as s on s.PK = i.SupplierKey where i.Status = 0  and s.SStatus = 0", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        Dim n As Integer = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Items :" & pb.Maximum & "/" & pb.Value

                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
                n = n + 1

                Dim strSQL As String = $"INSERT INTO tbl_Items 
                                    ([PK],
                                    ItemCode,
                                    ItemDescription,
                                    ItemType,
                                    ECRDescription,
                                    StockNumber,
                                    UnitOfMeasure,
                                    ClassKey,
                                    SupplierKey,
                                    [Discount],
                                    Commission,
                                    Terms,
                                    Remarks,
                                    ForeignCost,
                                    GrossCost,
                                    [Vat],
                                    [MarkUp],
                                    GrossSRP,
                                    LastModifiedBy,
                                    PhasedOut,
                                    BrandKey,
                                    ProdLineKey,
                                    OldCode,
                                    SeasonCode,
                                    [Change],
                                    MinQty,
                                    MaxQty,
                                    ReOrder,
                                    Category,
                                    PromoDisc,
                                    PromoDiscAmt,
                                    PromoFrom,
                                    PromoTo,
                                    PromoDiscLocked,
                                    Level1,
                                    Level2,
                                    Level3,
                                    Level4,
                                    Level5,
                                    Disc1,
                                    Disc2,
                                    Disc3,
                                    Disc4,
                                    Disc5,
                                    LastCost,
                                    LastSRP,
                                    [Color],
                                    StoreLocation,
                                    [PO],
                                    Date_Encoded,
                                    User_Action,
                                    User_Encoded,
                                    [Changes],
                                    RefNoID,
                                    NotIncludeInSale,
                                    [Active],
                                    ActiveAsOf,
                                    [Discounted],
                                    [MarkDown],
                                    [Status])
                                    VALUES ({rs.Fields("PK").Value},  
                                    '{fSqlFormat(rs.Fields("ItemCode").Value)}',
                                    '{fSqlFormat(rs.Fields("ItemDescription").Value)}',     
                                    {fNum(rs.Fields("ItemType").Value)},   
                                   '{fSqlFormat(rs.Fields("ECRDescription").Value)}',     
                                   '{fSqlFormat(rs.Fields("StockNumber").Value)}',     
                                   '{fSqlFormat(rs.Fields("UnitOfMeasure").Value)}',     
                                    {fNum(rs.Fields("ClassKey").Value)},   
                                    {fNum(rs.Fields("SupplierKey").Value)},   
                                    '{fSqlFormat(rs.Fields("Discount").Value)}', 
                                    '{fSqlFormat(rs.Fields("Commission").Value)}', 
                                    '{fSqlFormat(rs.Fields("Terms").Value)}', 
                                    '{fSqlFormat(rs.Fields("Remarks").Value)}',
                                    '{fSqlFormat(rs.Fields("ForeignCost").Value)}',
                                     {fNum(rs.Fields("GrossCost").Value)},  
                                     {fNum(rs.Fields("Vat").Value)},  
                                     {fNum(rs.Fields("MarkUp").Value)},  
                                     {fNum(rs.Fields("GrossSRP").Value)},  
                                    '{fSqlFormat(rs.Fields("LastModifiedBy").Value)}',
                                     {fNum(rs.Fields("PhasedOut").Value)},  
                                     {fNum(rs.Fields("BrandKey").Value)}, 
                                     {fNum(rs.Fields("ProdLineKey").Value)},  
                                     '{fSqlFormat(rs.Fields("OldCode").Value)}',
                                     '{fSqlFormat(rs.Fields("SeasonCode").Value)}',
                                     {fNum(rs.Fields("Change").Value)},  
                                     {fNum(rs.Fields("MinQty").Value)},  
                                     {fNum(rs.Fields("MaxQty").Value)},  
                                     {fNum(rs.Fields("ReOrder").Value)}, 
                                     {fNum(rs.Fields("Category").Value)}, 
                                     {fNum(rs.Fields("PromoDisc").Value)}, 
                                     {fNum(rs.Fields("PromoDiscAmt").Value)}, 
                                     {fDateIsEmpty(rs.Fields("PromoFrom").Value.ToString())},    
                                     {fDateIsEmpty(rs.Fields("PromoTo").Value.ToString())},   
                                     {fNum(rs.Fields("PromoDiscLocked").Value)},  
                                     {fNum(rs.Fields("Level1").Value)},  
                                     {fNum(rs.Fields("Level2").Value)},  
                                     {fNum(rs.Fields("Level3").Value)},  
                                     {fNum(rs.Fields("Level4").Value)},  
                                     {fNum(rs.Fields("Level5").Value)},  
                                     {fNum(rs.Fields("Disc1").Value)},  
                                     {fNum(rs.Fields("Disc2").Value)},  
                                     {fNum(rs.Fields("Disc3").Value)},  
                                     {fNum(rs.Fields("Disc4").Value)},  
                                     {fNum(rs.Fields("Disc5").Value)},  
                                     {fNum(rs.Fields("LastCost").Value)},
                                     {fNum(rs.Fields("LastSRP").Value)},
                                    '{fSqlFormat(rs.Fields("Color").Value)}',
                                     {fNum(rs.Fields("StoreLocation").Value)},
                                     {fNum(rs.Fields("PO").Value)},
                                     {fDateIsEmpty(rs.Fields("Date_Encoded").Value.ToString())},
                                     '{fSqlFormat(rs.Fields("User_Action").Value)}',   
                                     '{fSqlFormat(rs.Fields("User_Encoded").Value)}',   
                                     '{fSqlFormat(rs.Fields("Changes").Value)}',  
                                      {fNum(rs.Fields("RefNoID").Value)},
                                      {fNum(rs.Fields("NotIncludeInSale").Value)},
                                      {fNum(rs.Fields("Active").Value)},
                                      {fDateIsEmpty(rs.Fields("ActiveAsOf").Value.ToString())},
                                      {fNum(rs.Fields("Discounted").Value)},
                                      {fNum(rs.Fields("MarkDown").Value)},
                                      {fNum(rs.Fields("Status").Value)}

                                   );"

                conn.Execute(strSQL)
                rs.MoveNext()
            End While
        End If
    End Sub
    Public Sub CreateTable_tbl_Concession_PCR(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_Concession_PCR (
                                                PK INTEGER PRIMARY KEY,
                                                CtrlNo TEXT(12) NOT NULL,
                                                SeriesNo TEXT(7),
                                                YYear TEXT(4),
                                                EntryDate DATETIME NOT NULL,
                                                Reference TEXT(50) NOT NULL,
                                                EffectDuration TEXT(250),
                                                DiscPercent TEXT(3),
                                                DiscAmount CURRENCY,
                                                DiscFrom CURRENCY,
                                                DiscTo CURRENCY,
                                                Effect1 TEXT(22),
                                                Effect2 TEXT(22),
                                                Effect3 TEXT(22),
                                                Effect4 TEXT(22),
                                                Effect5 TEXT(22),
                                                Effect6 TEXT(22),
                                                Effect7 TEXT(22),
                                                Effect8 TEXT(22),
                                                Effect9 TEXT(22),
                                                Effect10 TEXT(22),
                                                Effect11 TEXT(22),
                                                Effect12 TEXT(22),
                                                Effect13 TEXT(22),
                                                Effect14 TEXT(22),
                                                Effect15 TEXT(22),
                                                SupplierKey INTEGER,
                                                SupplierCode TEXT(8),
                                                SupplierName TEXT(50),
                                                DeptKey INTEGER,
                                                DeptCode TEXT(3),
                                                DeptName TEXT(30),
                                                BrandKey INTEGER,
                                                BrandCode TEXT(5),
                                                BrandName TEXT(25),
                                                PCRType TEXT(1) NOT NULL,
                                                PreTag INTEGER NOT NULL,
                                                ForBarcode INTEGER NOT NULL,
                                                BarcodeType INTEGER NOT NULL,
                                                BarcodeColor INTEGER NOT NULL,
                                                PerSupplierBrand BYTE NOT NULL,
                                                PerBrand BYTE NOT NULL,
                                                PerPLU BYTE NOT NULL,
                                                Remarks TEXT(100) NOT NULL,
                                                TotalQty DOUBLE,
                                                TotalSRP DOUBLE,
                                                Disc TEXT(10),
                                                LastModifiedBy TEXT(50),
                                                Posted BYTE NOT NULL,
                                                PostedBy TEXT(50),
                                                PreparedBy TEXT(20),
                                                EncodedBy TEXT(20),
                                                CheckedBy TEXT(20),
                                                NotedBy TEXT(20),
                                                ApprovedBy TEXT(20),
                                                EffectTo1 TEXT(22),
                                                IsExtended BYTE NOT NULL,
                                                ExtendedBy TEXT(50),
                                                Unique_Effect INTEGER,
                                                CtrlNo_O TEXT(12),
                                                PL_Equation BYTE NOT NULL,
                                                PL_Amount CURRENCY NOT NULL,
                                                Sel_Reg BYTE NOT NULL,
                                                Sel_MD BYTE NOT NULL,
                                                Sel_PL BYTE NOT NULL
                                            );
"

            conn.Execute(createTableSql)
            Collect_tbl_Concession_PCR(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Concession_PCR")
            Application.Exit()
        End Try
    End Sub
    Private Sub Collect_tbl_Concession_PCR(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 1
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_Concession_PCR", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0

        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Concession_PCR :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim strSQL As String = $"INSERT INTO tbl_Concession_PCR 
                                    (PK,
                                    CtrlNo,
                                    SeriesNo,
                                    YYear,
                                    EntryDate,
                                    Reference,
                                    EffectDuration,
                                    DiscPercent,
                                    DiscAmount,
                                    DiscFrom,
                                    DiscTo,
                                    Effect1,
                                    Effect2,
                                    Effect3,
                                    Effect4,
                                    Effect5,
                                    Effect6,
                                    Effect7,
                                    Effect8,
                                    Effect9,
                                    Effect10,
                                    Effect11,
                                    Effect12,
                                    Effect13,
                                    Effect14,
                                    Effect15 ,
                                    SupplierKey,
                                    SupplierCode,
                                    SupplierName,
                                    DeptKey,
                                    DeptCode,
                                    DeptName,
                                    BrandKey,
                                    BrandCode,
                                    BrandName,
                                    PCRType,
                                    PreTag,
                                    ForBarcode,
                                    BarcodeType,
                                    BarcodeColor,
                                    PerSupplierBrand,
                                    PerBrand,
                                    PerPLU,
                                    Remarks,
                                    TotalQty,
                                    TotalSRP,
                                    Disc,
                                    LastModifiedBy,
                                    Posted,
                                    PostedBy,
                                    PreparedBy,
                                    EncodedBy,
                                    CheckedBy,
                                    NotedBy,
                                    ApprovedBy,
                                    EffectTo1,
                                    IsExtended,
                                    ExtendedBy,
                                    Unique_Effect,
                                    CtrlNo_O,
                                    PL_Equation,
                                    PL_Amount,
                                    Sel_Reg,
                                    Sel_MD,
                                    Sel_PL)
                                    VALUES ({rs.Fields("PK").Value},  
                                    '{fSqlFormat(rs.Fields("CtrlNo").Value)}',
                                    '{fSqlFormat(rs.Fields("SeriesNo").Value)}',
                                    '{fSqlFormat(rs.Fields("YYear").Value)}',     
                                     {fDateIsEmpty(rs.Fields("EntryDate").Value.ToString())},   
                                    '{fSqlFormat(rs.Fields("Reference").Value)}',  
                                    '{fSqlFormat(rs.Fields("EffectDuration").Value)}',       
                                    '{fSqlFormat(rs.Fields("DiscPercent").Value)}',   
                                     {fNum(rs.Fields("DiscAmount").Value)},  
                                    {fNum(rs.Fields("DiscFrom").Value)},  
                                    {fNum(rs.Fields("DiscTo").Value)}, 
                                    '{fSqlFormat(rs.Fields("Effect1").Value)}',  
                                    '{fSqlFormat(rs.Fields("Effect2").Value)}',  
                                    '{fSqlFormat(rs.Fields("Effect3").Value)}',  
                                     '{fSqlFormat(rs.Fields("Effect4").Value)}',  
                                     '{fSqlFormat(rs.Fields("Effect5").Value)}',  
                                     '{fSqlFormat(rs.Fields("Effect6").Value)}',  
                                     '{fSqlFormat(rs.Fields("Effect7").Value)}',  
                                     '{fSqlFormat(rs.Fields("Effect8").Value)}',  
                                     '{fSqlFormat(rs.Fields("Effect9").Value)}',  
                                     '{fSqlFormat(rs.Fields("Effect10").Value)}',  
                                     '{fSqlFormat(rs.Fields("Effect11").Value)}',  
                                     '{fSqlFormat(rs.Fields("Effect12").Value)}',  
                                     '{fSqlFormat(rs.Fields("Effect13").Value)}',  
                                     '{fSqlFormat(rs.Fields("Effect14").Value)}',  
                                     '{fSqlFormat(rs.Fields("Effect15").Value)}',  
                                     {fNum(rs.Fields("SupplierKey").Value)}, 
                                     '{fSqlFormat(rs.Fields("SupplierCode").Value)}',  
                                     '{fSqlFormat(rs.Fields("SupplierName").Value)}', 
                                     {fNum(rs.Fields("DeptKey").Value)},                                         
                                     '{fSqlFormat(rs.Fields("DeptCode").Value)}',  
                                     '{fSqlFormat(rs.Fields("DeptName").Value)}',
                                     {fNum(rs.Fields("BrandKey").Value)},  
                                     '{fSqlFormat(rs.Fields("BrandCode").Value)}',
                                     '{fSqlFormat(rs.Fields("BrandName").Value)}',
                                     '{fSqlFormat(rs.Fields("PCRType").Value)}',
                                      {fNum(rs.Fields("PreTag").Value)},  
                                      {fNum(rs.Fields("ForBarcode").Value)},  
                                      {fNum(rs.Fields("BarcodeType").Value)},  
                                      {fNum(rs.Fields("BarcodeColor").Value)},  
                                      {fNum(rs.Fields("PerSupplierBrand").Value)},  
                                      {fNum(rs.Fields("PerBrand").Value)},  
                                      {fNum(rs.Fields("PerPLU").Value)},  
                                     '{fSqlFormat(rs.Fields("Remarks").Value)}',
                                      {fNum(rs.Fields("TotalQty").Value)},  
                                      {fNum(rs.Fields("TotalSRP").Value)},  
                                     '{fSqlFormat(rs.Fields("Disc").Value)}',
                                     '{fSqlFormat(rs.Fields("LastModifiedBy").Value)}',  
                                      {fNum(rs.Fields("Posted").Value)},  
                                      '{fSqlFormat(rs.Fields("PostedBy").Value)}',  
                                      '{fSqlFormat(rs.Fields("PreparedBy").Value)}',  
                                      '{fSqlFormat(rs.Fields("EncodedBy").Value)}', 
                                      '{fSqlFormat(rs.Fields("CheckedBy").Value)}',   
                                      '{fSqlFormat(rs.Fields("NotedBy").Value)}',  
                                      '{fSqlFormat(rs.Fields("ApprovedBy").Value)}',  
                                      '{fSqlFormat(rs.Fields("EffectTo1").Value)}',  
                                       {fNum(rs.Fields("IsExtended").Value)}, 
                                      '{fSqlFormat(rs.Fields("ExtendedBy").Value)}',  
                                       {fNum(rs.Fields("Unique_Effect").Value)}, 
                                      '{fSqlFormat(rs.Fields("CtrlNo_O").Value)}',  
                                       {fNum(rs.Fields("PL_Equation").Value)}, 
                                       {fNum(rs.Fields("PL_Amount").Value)}, 
                                       {fNum(rs.Fields("Sel_Reg").Value)}, 
                                       {fNum(rs.Fields("Sel_MD").Value)}, 
                                       {fNum(rs.Fields("Sel_PL").Value)}
                                   );"

                conn.Execute(strSQL)
                rs.MoveNext()


            End While
        End If
    End Sub


    Public Sub CreateTable_tbl_Concession_PCR_Det(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = " CREATE TABLE tbl_Concession_PCR_Det (
                                                PK INTEGER PRIMARY KEY,
                                                ConcPCRKey INTEGER NOT NULL,
                                                Line INTEGER,
                                                ItemKey INTEGER NOT NULL,
                                                Qty DOUBLE NOT NULL,
                                                SRP DOUBLE NOT NULL,
                                                S_SRP DOUBLE NOT NULL,
                                                O_Remarks TEXT(15) NOT NULL,
                                                Posted BYTE NOT NULL,
                                                BarcodeQty DOUBLE NOT NULL,
                                                RevisedPLU TEXT(12),
                                                DiscPercent_det TEXT(3),
                                                DiscAmount_det CURRENCY,
                                                DiscNewSRP CURRENCY,
                                                Duration_ByItem TEXT(200),
                                                TotalSRP CURRENCY,
                                                O_SRP CURRENCY NOT NULL,
                                                Remarks TEXT(100),
                                                SupplierKey INTEGER,
                                                BrandKey INTEGER,
                                                Selected BYTE NOT NULL,
                                                StockNo TEXT(25) NOT NULL,
                                                RefCtrlNo TEXT(15) NOT NULL,
                                                RefConcPCRKey INTEGER NOT NULL,
                                                BaseSRP_new DOUBLE NOT NULL,
                                                DiscountedSRP_new DOUBLE NOT NULL,
                                                DiscPercent_new TEXT(3),
                                                DiscAmount_new TEXT(3),
                                                BrandName TEXT(30) NOT NULL,
                                                IsCurrentlyMarkdown BYTE NOT NULL
);"

            conn.Execute(createTableSql)
            Collect_tbl_Concession_PCR_Det(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Concession_PCR_Det")
            Application.Exit()
        End Try
    End Sub

    Private Sub Collect_tbl_Concession_PCR_Det(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 1
        rs = New ADODB.Recordset
        rs.Open($"select dd.* from [tbl_Concession_PCR_Det] as dd INNER JOIN tbl_Concession_PCR on tbl_Concession_PCR.PK = dd.ConcPCRKey ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0

        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Concession_PCR_Det :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim strSQL As String = $"INSERT INTO tbl_Concession_PCR_Det 
                                                (PK,
                                                ConcPCRKey,
                                                Line,
                                                ItemKey,
                                                Qty,
                                                SRP,
                                                S_SRP,
                                                O_Remarks,
                                                Posted,
                                                BarcodeQty,
                                                RevisedPLU,
                                                DiscPercent_det,
                                                DiscAmount_det,
                                                DiscNewSRP,
                                                Duration_ByItem,
                                                TotalSRP,
                                                O_SRP,
                                                Remarks,
                                                SupplierKey,
                                                BrandKey,
                                                Selected,
                                                StockNo,
                                                RefCtrlNo,
                                                RefConcPCRKey,
                                                BaseSRP_new,
                                                DiscountedSRP_new,
                                                DiscPercent_new,
                                                DiscAmount_new,
                                                BrandName,
                                                IsCurrentlyMarkdown)
                                    VALUES ({fNum(rs.Fields("PK").Value)},  
                                    {fNum(rs.Fields("ConcPCRKey").Value)},
                                    {fNum(rs.Fields("Line").Value)},
                                    {fNum(rs.Fields("ItemKey").Value)},
                                    {fNum(rs.Fields("Qty").Value)},     
                                    {fNum(rs.Fields("SRP").Value)},   
                                    {fNum(rs.Fields("S_SRP").Value)},
                                   '{fSqlFormat(rs.Fields("O_Remarks").Value)}',   
                                    {fNum(rs.Fields("Posted").Value)},   
                                    {fNum(rs.Fields("BarcodeQty").Value)},     
                                   '{fSqlFormat(rs.Fields("RevisedPLU").Value)}',   
                                   '{fSqlFormat(rs.Fields("DiscPercent_det").Value)}', 
                                    {fNum(rs.Fields("DiscAmount_det").Value)}, 
                                    {fNum(rs.Fields("DiscNewSRP").Value)}, 
                                   '{fSqlFormat(rs.Fields("Duration_ByItem").Value)}',  
                                    {fNum(rs.Fields("TotalSRP").Value)},
                                    {fNum(rs.Fields("O_SRP").Value)},
                                   '{fSqlFormat(rs.Fields("Remarks").Value)}',   
                                    {fNum(rs.Fields("SupplierKey").Value)},
                                    {fNum(rs.Fields("BrandKey").Value)},
                                    {fNum(rs.Fields("Selected").Value)},
                                   '{fSqlFormat(rs.Fields("StockNo").Value)}', 
                                   '{fSqlFormat(rs.Fields("RefCtrlNo").Value)}', 
                                    {fNum(rs.Fields("RefConcPCRKey").Value)},
                                    {fNum(rs.Fields("BaseSRP_new").Value)},
                                    {fNum(rs.Fields("DiscountedSRP_new").Value)},
                                   '{fSqlFormat(rs.Fields("DiscPercent_new").Value)}', 
                                   '{fSqlFormat(rs.Fields("DiscAmount_new").Value)}', 
                                   '{fSqlFormat(rs.Fields("BrandName").Value)}', 
                                    {fNum(rs.Fields("IsCurrentlyMarkdown").Value)}
                                    
                                   );"

                conn.Execute(strSQL)
                rs.MoveNext()


            End While
        End If
    End Sub



    Public Sub CreateTable_tbl_Concession_PCR_Effectivity(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_Concession_PCR_Effectivity (
                                                    PK INTEGER PRIMARY KEY,
                                                    ConcPCRKey INTEGER NOT NULL,
                                                    Effect_From DATETIME,
                                                    Effect_To DATETIME,
                                                    Posted BYTE,
                                                    IsExtended BYTE,
                                                    ExtendedBy TEXT(50),
                                                    LastModifiedBy TEXT(100) NOT NULL
                                                );"

            conn.Execute(createTableSql)
            Collect_tbl_Concession_PCR_Effectivity(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_Concession_PCR_Effectivity")
            Application.Exit()
        End Try
    End Sub

    Private Sub Collect_tbl_Concession_PCR_Effectivity(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 1
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_Concession_PCR_Effectivity", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0

        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Concession_PCR_Effectivity :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim strSQL As String = $"INSERT INTO tbl_Concession_PCR_Effectivity 
                                    (PK,
                                    ConcPCRKey,
                                    Effect_From,
                                    Effect_To,
                                    Posted,
                                    IsExtended,
                                    ExtendedBy,
                                    LastModifiedBy)
                                    VALUES ({fNum(rs.Fields("PK").Value)}, 
                                    {fNum(rs.Fields("ConcPCRKey").Value)}, 
                                    {fDateIsEmpty(rs.Fields("Effect_From").Value.ToString())},
                                    {fDateIsEmpty(rs.Fields("Effect_To").Value.ToString())},
                                    {fNum(rs.Fields("Posted").Value)}, 
                                    {fNum(rs.Fields("IsExtended").Value)},    
                                    {fNum(rs.Fields("ExtendedBy").Value)},   
                                    {fNum(rs.Fields("LastModifiedBy").Value)}                   
                                   );"

                conn.Execute(strSQL)
                rs.MoveNext()

            End While
        End If
    End Sub



    Public Sub CreateTable_tbl_GiftCert_Changes(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_GiftCert_Changes (
                                            PK INTEGER PRIMARY KEY,
                                            EffectDate DATETIME NOT NULL,
                                            GCNumber DOUBLE NOT NULL,
                                            GCAmount DOUBLE NOT NULL,
                                            Changes TEXT(50) NOT NULL
                                        );"

            conn.Execute(createTableSql)
            Collect_tbl_GiftCert_Changes(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_GiftCert_Changes")
            Application.Exit()
        End Try
    End Sub

    Private Sub Collect_tbl_GiftCert_Changes(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 1

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_GiftCert_Changes ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_GiftCert_Changes :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim strSQL As String = $"INSERT INTO tbl_GiftCert_Changes 
                                    (PK,
                                    EffectDate,
                                    GCNumber,
                                    GCAmount,
                                    [Changes])
                                    VALUES ({rs.Fields("PK").Value},     
                                    {fDateIsEmpty(rs.Fields("EffectDate").Value.ToString())}, 
                                    {rs.Fields("GCNumber").Value},
                                    {rs.Fields("GCAmount").Value},
                                   '{fSqlFormat(rs.Fields("Changes").Value)}'
                                );"
                conn.Execute(strSQL)
                rs.MoveNext()
            End While

        End If

    End Sub
    Public Sub CreateTable_tbl_PS_Upload_Utility(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_PS_Upload_Utility (
                            EffectDate DATETIME NOT NULL,
                            StopUpload BYTE NOT NULL
                        );"

            conn.Execute(createTableSql)
            Collect_tbl_PS_Upload_Utility(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_GiftCert_Changes")
            Application.Exit()
        End Try
    End Sub

    Private Sub Collect_tbl_PS_Upload_Utility(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 1

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PS_Upload_Utility ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_Upload_Utility :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim strSQL As String = $"INSERT INTO tbl_PS_Upload_Utility 
                                    (EffectDate,
                                    StopUpload)
                                    VALUES (    
                                    {fDateIsEmpty(rs.Fields("EffectDate").Value.ToString())}, 
                                    {rs.Fields("StopUpload").Value}
                                );"
                conn.Execute(strSQL)
                rs.MoveNext()
            End While

        End If

    End Sub

    Public Sub CreateTable_tbl_VPlus_Codes_Changes(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_VPlus_Codes_Changes (
                                            Codes TEXT(16) NOT NULL,
                                            DateChange DATETIME NOT NULL
                                        );"

            conn.Execute(createTableSql)
            Collect_tbl_VPlus_Codes_Changes(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_VPlus_Codes_Changes")
            Application.Exit()
        End Try
    End Sub
    Private Sub Collect_tbl_VPlus_Codes_Changes(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 5

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_VPlus_Codes_Changes ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_VPlus_Codes_Changes :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim strSQL As String = $"INSERT INTO tbl_VPlus_Codes_Changes 
                                    (Codes,
                                    DateChange)
                                    VALUES ('{fSqlFormat(rs.Fields("Codes").Value)}',               
                                    {fDateIsEmpty(rs.Fields("DateChange").Value.ToString())}   
                                   );"

                conn.Execute(strSQL)
                rs.MoveNext()
            End While

        End If

    End Sub
    Public Sub CreateTable_tbl_VPlus_Summary(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_VPlus_Summary (
                                            PK INTEGER PRIMARY KEY,
                                            VPlusCode TEXT(16) NOT NULL,
                                            TransDate DATETIME NOT NULL,
                                            Location TEXT(1) NOT NULL,
                                            Cash DOUBLE NOT NULL,
                                            Card DOUBLE NOT NULL,
                                            [GC] DOUBLE NOT NULL,
                                            VPlus DOUBLE NOT NULL,
                                            InOut TEXT(1) NOT NULL,
                                            InPoints DOUBLE NOT NULL,
                                            OutPoints DOUBLE NOT NULL
                                        );"

            conn.Execute(createTableSql)
            Collect_tbl_VPlus_Summary(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_VPlus_Summary")
            Application.Exit()
        End Try
    End Sub
    Private Sub Collect_tbl_VPlus_Summary(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 5

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_VPlus_Summary ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_VPlus_Summary :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim strSQL As String = $"INSERT INTO tbl_VPlus_Summary 
                                            (PK,
                                            VPlusCode,
                                            TransDate,
                                            Location,
                                            Cash,
                                            Card,
                                            [GC],
                                            VPlus,
                                            InOut ,
                                            InPoints,
                                            OutPoints)
                                    VALUES ({fNum(rs.Fields("PK").Value)},
                                        '{fSqlFormat(rs.Fields("VPlusCode").Value)}',               
                                         {fDateIsEmpty(rs.Fields("TransDate").Value.ToString())},
                                        '{fSqlFormat(rs.Fields("Location").Value)}', 
                                         {fNum(rs.Fields("Cash").Value)},  
                                         {fNum(rs.Fields("Card").Value)},  
                                         {fNum(rs.Fields("GC").Value)},    
                                         {fNum(rs.Fields("VPlus").Value)},  
                                        '{fSqlFormat(rs.Fields("InOut").Value)}', 
                                        {fNum(rs.Fields("InPoints").Value)},  
                                        {fNum(rs.Fields("OutPoints").Value)} 
                                   );"

                conn.Execute(strSQL)
                rs.MoveNext()
            End While

        End If

    End Sub
    Public Sub CreateTable_tbl_VPlus_Codes_For_Offline(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_VPlus_Codes_For_Offline (
                                            Codes TEXT(16) NOT NULL,
                                            POSName TEXT(3) NOT NULL,
                                            Used BYTE NOT NULL,
                                            CreatedOn DATETIME NOT NULL
                                        );"

            conn.Execute(createTableSql)
            Collect_tbl_VPlus_Codes_For_Offline(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_VPlus_Codes_For_Offline")
            Application.Exit()
        End Try
    End Sub
    Private Sub Collect_tbl_VPlus_Codes_For_Offline(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 5

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_VPlus_Codes_For_Offline ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_VPlus_Codes_For_Offline :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim strSQL As String = $"INSERT INTO tbl_VPlus_Codes_For_Offline 
                                            (Codes,
                                            POSName,
                                            Used,
                                            CreatedOn)
                                    VALUES (
                                        '{fSqlFormat(rs.Fields("Codes").Value)}',    
                                        '{fSqlFormat(rs.Fields("POSName").Value)}',    
                                         {fNum(rs.Fields("Used").Value)},          
                                         {fDateIsEmpty(rs.Fields("CreatedOn").Value.ToString())}
                                   );"

                conn.Execute(strSQL)
                rs.MoveNext()
            End While

        End If

    End Sub
    Public Sub CreateTable_tbl_VPlus_App(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_VPlus_App (
                                            PLU TEXT(12) Not NULL
                                        );"

            conn.Execute(createTableSql)
            Collect_tbl_VPlus_App(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_VPlus_App")
            Application.Exit()
        End Try
    End Sub
    Private Sub Collect_tbl_VPlus_App(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 5

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_VPlus_App ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_VPlus_App :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim strSQL As String = $"INSERT INTO tbl_VPlus_App 
                                            (PLU)
                                    VALUES (
                                        '{fSqlFormat(rs.Fields("PLU").Value)}'
                                   );"

                conn.Execute(strSQL)
                rs.MoveNext()
            End While

        End If

    End Sub
    Public Sub CreateTable_tbl_RetrieveHistoryForLocal(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_RetrieveHistoryForLocal (
                                                Counter TEXT(50) NOT NULL,
                                                ForRetrieval DECIMAL(18, 0) NOT NULL
                                            );"

            conn.Execute(createTableSql)
            Collect_tbl_RetrieveHistoryForLocal(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_VPlus_App")
            Application.Exit()
        End Try
    End Sub
    Private Sub Collect_tbl_RetrieveHistoryForLocal(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 5

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_RetrieveHistoryForLocal ", ConnMain, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_RetrieveHistoryForLocal :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim strSQL As String = $"INSERT INTO tbl_RetrieveHistoryForLocal 
                                            (Counter,
                                            ForRetrieval)
                                    VALUES (
                                        '{fSqlFormat(rs.Fields("Counter").Value)}',
                                        {fNum(rs.Fields("ForRetrieval").Value)}
                                   );"

                conn.Execute(strSQL)
                rs.MoveNext()
            End While

        End If

    End Sub

End Module
