Imports ADODB
Module ModMainImport


    Public Sub Insert_tbl_PCPOS_Cashiers(pb As ProgressBar, l As Label)
        ConnServer.Execute("SET IDENTITY_INSERT tbl_PCPOS_Cashiers ON;")
        rs = New ADODB.Recordset
        rs.Open("select * from tbl_PCPOS_Cashiers ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then

            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PCPOS_Cashiers :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim rx As New ADODB.Recordset
                rx.Open($"select top 1 * from tbl_PCPOS_Cashiers where CashierCode = '{rs.Fields("CashierCode").Value}' ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
                    ConnServer.Execute($"INSERT INTO tbl_PCPOS_Cashiers 
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
                End If


                rs.MoveNext()
            End While
        End If
        ConnServer.Execute("SET IDENTITY_INSERT tbl_PCPOS_Cashiers OFF;")
    End Sub
    Public Sub Insert_tbl_ItemsForPLU(pb As ProgressBar, l As Label)
        rs = New ADODB.Recordset
        rs.Open("select tbl_ItemsForPLU.*  FROM tbl_ItemsForPLU  ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
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
                Dim rx As New ADODB.Recordset
                rx.Open($"select top 1 tbl_ItemsForPLU.*  FROM tbl_ItemsForPLU  where ItemCode='{rs.Fields("ItemCode").Value}' ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
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
                                     {fNum(rs.Fields("GrossSRP").Value)},
                                     {fNum(rs.Fields("PromoDisc").Value)},
                                     {fDateIsEmpty(rs.Fields("PromoFrom").Value.ToString())},
                                     {fDateIsEmpty(rs.Fields("PromoTo").Value.ToString())}
                                );"
                    ConnServer.Execute(strSQL)
                End If
                rs.MoveNext()
            End While
        End If
    End Sub
    Public Sub Insert_tbl_Bank(pb As ProgressBar, l As Label)
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Bank ON;")
        rs = New ADODB.Recordset
        rs.Open("select * from tbl_Bank ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Bank :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim rx As New ADODB.Recordset
                rx.Open($"SELECT TOP 1 * FROM tbl_bank WHERE PK = {rs.Fields("PK").Value} ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_Bank 
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
                    ConnServer.Execute(strSQL)
                End If
                rs.MoveNext()
            End While

        End If
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Bank OFF;")
    End Sub
    Public Sub Insert_tbl_Banks(pb As ProgressBar, l As Label)
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Banks ON;")
        rs = New ADODB.Recordset
        rs.Open("select * from tbl_Banks ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Banks :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim rx As New ADODB.Recordset()
                rx.Open($"select TOP 1 * from tbl_Banks  WHERE PK = { fNum(rs.Fields("PK").Value)}", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_Banks 
                                            (PK,
                                            BankCode,
                                            BankName,
                                            Telephone,
                                            MERC_COD,
                                            MERC_COD2,
                                            [Description],
                                            Bank)
                                            VALUES ({ fNum(rs.Fields("PK").Value)},
                                            '{fSqlFormat(rs.Fields("BankCode").Value)}',
                                            '{fSqlFormat(rs.Fields("BankName").Value)}',
                                            '{fSqlFormat(rs.Fields("Telephone").Value)}',
                                            '{fSqlFormat(rs.Fields("MERC_COD").Value)}',
                                            '{fSqlFormat(rs.Fields("MERC_COD2").Value)}',
                                            '{fSqlFormat(rs.Fields("Description").Value)}',
                                             {fNum(rs.Fields("Bank").Value)}       
                                             );"
                    ConnServer.Execute(strSQL)
                End If
                rs.MoveNext()
            End While

        End If
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Banks OFF;")

    End Sub
    Public Sub Insert_tbl_Banks_Changes(pb As ProgressBar, l As Label)

        ConnServer.Execute("SET IDENTITY_INSERT tbl_Banks_Changes ON;")
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_Banks_Changes ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Banks_Changes :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()

                Dim rx As New ADODB.Recordset
                rx.Open($"select TOP 1 * from tbl_Banks_Changes WHERE PK = {fNum(rs.Fields("PK").Value)}", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_Banks_Changes 
                                    (PK,
                                    EffectDate,
                                    BankKey,
                                    [Changes])
                                    VALUES ({fNum(rs.Fields("PK").Value)},  
                                    {fDateIsEmpty(rs.Fields("EffectDate").Value.ToString())},
                                    {fNum(rs.Fields("BankKey").Value)},    
                                    '{fSqlFormat(rs.Fields("Changes").Value)}'   
                                   );"

                    ConnServer.Execute(strSQL)
                End If
                rs.MoveNext()
            End While

        End If
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Banks_Changes OFF;")
    End Sub
    Public Sub Insert_tbl_Bank_Changes(pb As ProgressBar, l As Label)
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Bank_Changes ON;")
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_Bank_Changes", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Bank_Changes :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim rx As New ADODB.Recordset
                rx.Open($"SELECT TOP 1 * FROM tbl_Bank_Changes WHERE PK = {fNum(rs.Fields("PK").Value)} ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_Bank_Changes 
                                    (PK,
                                    EffectDate,
                                    BankKey,
                                    [Changes])
                                    VALUES ({fNum(rs.Fields("PK").Value)},  
                                    {fDateIsEmpty(rs.Fields("EffectDate").Value.ToString())},
                                    {fNum(rs.Fields("BankKey").Value)},    
                                    '{fSqlFormat(rs.Fields("Changes").Value)}'   
                                   );"

                    ConnServer.Execute(strSQL)
                End If
                rs.MoveNext()
            End While

        End If
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Bank_Changes OFF;")
    End Sub
    Public Sub Insert_tbl_Bank_Terms(pb As ProgressBar, l As Label)


        rs = New Recordset
        rs.Open("select * from tbl_Bank_Terms ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Bank_Terms :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim rx As New Recordset
                rx.Open($"select TOP 1 * from tbl_Bank_Terms where BankKey={rs.Fields("BankKey").Value} and Effectivity = {fDateIsEmpty(rs.Fields("Effectivity").Value.ToString())} ", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
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
                    ConnServer.Execute(strSQL)
                End If

                rs.MoveNext()
            End While

        End If


    End Sub
    Public Sub Insert_tbl_QRPay_Type(pb As ProgressBar, l As Label)
        ConnServer.Execute("SET IDENTITY_INSERT tbl_QRPay_Type ON;")

        rs = New ADODB.Recordset
        rs.Open("select * from tbl_QRPay_Type ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_QRPay_Type :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim rx As New Recordset
                rx.Open($"select TOP 1 * from tbl_QRPay_Type WHERE nQRPTypeID = { fNum(rs.Fields("nQRPTypeID").Value)} ", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
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
                    ConnServer.Execute(strSQL)
                End If
                rs.MoveNext()
            End While
        End If
        ConnServer.Execute("SET IDENTITY_INSERT tbl_QRPay_Type OFF;")
    End Sub
    Public Sub Insert_tbl_GiftCert_List(pb As ProgressBar, l As Label)
        ConnServer.Execute("SET IDENTITY_INSERT tbl_GiftCert_List ON;")

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_GiftCert_List ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_GiftCert_List :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim rx As New Recordset
                rx.Open($"select * from tbl_GiftCert_List where PK={rs.Fields("PK").Value} ", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
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
                    ConnServer.Execute(strSQL)
                End If

                rs.MoveNext()
            End While

        End If
        ConnServer.Execute("SET IDENTITY_INSERT tbl_GiftCert_List OFF;")
    End Sub
    Public Sub Insert_tbl_VPlus_Codes(pb As ProgressBar, l As Label)

        Dim n As Integer = 0
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_VPlus_Codes ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_VPlus_Codes :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If

                Dim rx As New Recordset
                rx.Open($"select TOP 1 * from tbl_VPlus_Codes where Codes = '{fSqlFormat(rs.Fields("Codes").Value)}' ", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_VPlus_Codes 
                                    (Codes,
                                    Customer,
                                    InPoints,
                                    OutPoints,                            
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
                                    {rs.Fields("Blocked").Value},
                                    {rs.Fields("Printed").Value},
                                    {fDateIsEmpty(rs.Fields("CreatedOn").Value.ToString())},
                                    {fDateIsEmpty(rs.Fields("CreatedOnTime").Value.ToString())},
                                   '{fSqlFormat(rs.Fields("Password").Value.ToString())}',
                                    {fDateIsEmpty(rs.Fields("DateStarted").Value.ToString())},
                                    {fDateIsEmpty(rs.Fields("DateExpired").Value.ToString())},
                                    {fDateIsEmpty(rs.Fields("DateModified").Value.ToString())},
                                    {rs.Fields("Changes").Value} );"
                    ConnServer.Execute(strSQL)
                End If

                rs.MoveNext()
            End While
        End If

    End Sub
    Public Sub Insert_tbl_VPlus_Codes_Validity(pb As ProgressBar, l As Label)


        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_VPlus_Codes_Validity ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        Dim n As Integer = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_VPlus_Codes_Validity :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If

                Dim rx As New Recordset
                rx.Open($"select TOP 1 * from tbl_VPlus_Codes_Validity where Codes = '{fSqlFormat(rs.Fields("Codes").Value)}'", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
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

                    ConnServer.Execute(strSQL)
                End If
                rs.MoveNext()
            End While

        End If

    End Sub
    Public Sub Insert_tbl_PCPOS_Cashiers_Changes(pb As ProgressBar, l As Label)

        ConnServer.Execute("SET IDENTITY_INSERT tbl_PCPOS_Cashiers_Changes ON;")
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PCPOS_Cashiers_Changes", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PCPOS_Cashiers_Changes :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim rx As New Recordset
                rx.Open($"select TOP 1 * from tbl_PCPOS_Cashiers_Changes WHERE PK = {fNum(rs.Fields("PK").Value)}  ", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
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

                    ConnServer.Execute(strSQL)
                End If
                rs.MoveNext()
            End While
        End If
        ConnServer.Execute("SET IDENTITY_INSERT tbl_PCPOS_Cashiers_Changes OFF;")
    End Sub
    Public Sub Insert_tbl_Items_Changes(pb As ProgressBar, l As Label)
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Items_Change ON;")
        rs = New ADODB.Recordset
        rs.Open($"select  * from tbl_Items_Change ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Items_Change :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()

                Dim rx As New Recordset
                rx.Open($"select TOP 1  * from tbl_Items_Change WHERE PK ={rs.Fields("PK").Value} ", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
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

                    ConnServer.Execute(strSQL)
                End If

                rs.MoveNext()
            End While
        End If
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Items_Change OFF;")
    End Sub

    Public Sub Insert_tbl_ItemsForPLU_For_Effect(pb As ProgressBar, l As Label)

        ConnServer.Execute("SET IDENTITY_INSERT tbl_ItemsForPLU_For_Effect ON;")
        Dim year As Integer = Now.Year - 1
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_ItemsForPLU_For_Effect", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0

        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_ItemsForPLU_For_Effect :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()

                Dim rx As New Recordset
                rx.Open($"SELECT TOP 1 * FROM tbl_ItemsForPLU_For_Effect WHERE PK = {rs.Fields("PK").Value}", ConnServer, CursorTypeEnum.adOpenStatic)
                If (rx.RecordCount = 0) Then
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
                    ConnServer.Execute(strSQL)
                End If
                rs.MoveNext()
            End While
        End If
        ConnServer.Execute("SET IDENTITY_INSERT tbl_ItemsForPLU_For_Effect OFF;")
    End Sub
    Public Sub Insert_tbl_Items(pb As ProgressBar, l As Label)
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Items ON;")

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_Items as i ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
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
                Dim rx As New Recordset
                rx.Open($"SELECT TOP 1 * FROM tbl_Items as i WHERE i.PK = {rs.Fields("PK").Value}", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
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
                    ConnServer.Execute(strSQL)
                End If
                rs.MoveNext()
            End While
        End If
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Items OFF;")
    End Sub

    Public Sub Insert_tbl_VPlus_Codes_Changes(pb As ProgressBar, l As Label)

        ConnServer.Execute("TRUNCATE TABLE tbl_VPlus_Codes_Changes;") ' clean table


        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_VPlus_Codes_Changes ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        Dim n As Integer = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_VPlus_Codes_Changes :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
                Dim rx As New Recordset

                Dim strSQL As String = $"INSERT INTO tbl_VPlus_Codes_Changes 
                    (Codes,
                    DateChange)
                    VALUES ('{fSqlFormat(rs.Fields("Codes").Value)}',               
                    {fDateIsEmpty(rs.Fields("DateChange").Value.ToString())}   
                    );"

                ConnServer.Execute(strSQL)


                rs.MoveNext()
            End While

        End If

    End Sub


    Public Sub Insert_tbl_Concession_PCR(pb As ProgressBar, l As Label)
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Concession_PCR ON;")
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_Concession_PCR  ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        Dim n As Integer = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Concession_PCR :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If

                Dim rx As New Recordset
                rx.Open($"select TOP 1 * from tbl_Concession_PCR WHERE PK ={rs.Fields("PK").Value} ", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
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

                    ConnServer.Execute(strSQL)
                End If

                rs.MoveNext()


            End While
        End If
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Concession_PCR OFF;")
    End Sub
    Public Sub Insert_tbl_Concession_PCR_Det(pb As ProgressBar, l As Label)
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Concession_PCR_Det ON;")
        rs = New ADODB.Recordset
        rs.Open($"select dd.* from [tbl_Concession_PCR_Det] as dd ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        Dim n As Integer = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Concession_PCR_Det :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
                Dim rx As New Recordset
                rx.Open($"select TOP 1 dd.* from [tbl_Concession_PCR_Det] as dd  where dd.PK = {fNum(rs.Fields("PK").Value)} ", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
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

                    ConnServer.Execute(strSQL)
                End If

                rs.MoveNext()
            End While
        End If
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Concession_PCR_Det OFF;")
    End Sub
    Public Sub Insert_tbl_Concession_PCR_Effectivity(pb As ProgressBar, l As Label)
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Concession_PCR_Effectivity ON;")
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_Concession_PCR_Effectivity ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        Dim n As Integer = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_Concession_PCR_Effectivity :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If

                Dim rx As New Recordset
                rx.Open($"select TOP 1 * from tbl_Concession_PCR_Effectivity WHERE PK = {fNum(rs.Fields("PK").Value)}", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
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

                    ConnServer.Execute(strSQL)
                End If

                rs.MoveNext()

            End While
        End If
        ConnServer.Execute("SET IDENTITY_INSERT tbl_Concession_PCR_Effectivity OFF;")
    End Sub

    Public Sub Insert_tbl_GiftCert_Changes(pb As ProgressBar, l As Label)

        ConnServer.Execute("SET IDENTITY_INSERT tbl_GiftCert_Changes ON;")
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_GiftCert_Changes ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_GiftCert_Changes :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim rx As New Recordset
                rx.Open($"select TOP 1 * from tbl_GiftCert_Change Where PK = {rs.Fields("PK").Value}", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount Then
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
                    ConnServer.Execute(strSQL)
                End If

                rs.MoveNext()
            End While

        End If

        ConnServer.Execute("SET IDENTITY_INSERT tbl_GiftCert_Changes OFF;")
    End Sub

    Public Sub Insert_tbl_PS_Upload_Utility(pb As ProgressBar, l As Label)

        ConnServer.Execute("TRUNCATE TABLE tbl_PS_Upload_Utility;") ' clean table
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PS_Upload_Utility ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
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
                ConnServer.Execute(strSQL)
                rs.MoveNext()
            End While
        End If

    End Sub

    Public Sub Insert_tbl_VPlus_Summary(pb As ProgressBar, l As Label)

        ConnServer.Execute("SET IDENTITY_INSERT tbl_VPlus_Summary ON;")
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_VPlus_Summary ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        Dim n As Integer = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_VPlus_Summary :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
                Dim rx As New Recordset
                rx.Open($"select TOP 1 * from tbl_VPlus_Summary WHERE PK = {fNum(rs.Fields("PK").Value)}", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
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
                    ConnServer.Execute(strSQL)
                End If
                rs.MoveNext()
            End While

        End If
        ConnServer.Execute("SET IDENTITY_INSERT tbl_VPlus_Summary OFF;")
    End Sub

    Public Sub Insert_tbl_VPlus_Codes_For_Offline(pb As ProgressBar, l As Label)
        ConnServer.Execute("TRUNCATE TABLE tbl_VPlus_Codes_For_Offline;") ' clean table
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_VPlus_Codes_For_Offline ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        Dim n As Integer = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_VPlus_Codes_For_Offline :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
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

                ConnServer.Execute(strSQL)
                rs.MoveNext()
            End While

        End If

    End Sub

    Public Sub Insert_tbl_VPlus_App(pb As ProgressBar, l As Label)

        ConnServer.Execute("TRUNCATE TABLE tbl_VPlus_App;") ' clean table

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_VPlus_App ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
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

                ConnServer.Execute(strSQL)
                rs.MoveNext()
            End While

        End If

    End Sub
    Public Sub Insert_tbl_RetrieveHistoryForLocal(pb As ProgressBar, l As Label)
        ConnServer.Execute("TRUNCATE TABLE tbl_RetrieveHistoryForLocal;") ' clean table

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_RetrieveHistoryForLocal ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_RetrieveHistoryForLocal :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim strSQL As String = $"INSERT INTO tbl_RetrieveHistoryForLocal 
                                            ([Counter],
                                            [ForRetrieval])
                                    VALUES (
                                        '{fSqlFormat(rs.Fields("Counter").Value)}',
                                        {fNum(rs.Fields("ForRetrieval").Value)}
                                   );"

                ConnServer.Execute(strSQL)
                rs.MoveNext()
            End While

        End If

    End Sub

    Public Sub Insert_tbl_PS_GT(pb As ProgressBar, l As Label)

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PS_GT ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_GT :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim rx As New Recordset
                rx.Open($"select * from tbl_PS_GT WHERE [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}'", ConnServer, CursorTypeEnum.adOpenStatic)
                If (rx.RecordCount = 0) Then
                    Dim strSQL As String = $"INSERT INTO tbl_PS_GT 
                                            ( [Counter],
                                                TransactionCount,
                                                GrandTotal,
                                                ZZCount,
                                                ResetCnt,
                                                ResetTrans,
                                                InvoiceNumberOld,
                                                InvoiceNumberCnt,
                                                InvoiceNumber,
                                                RA,
                                                RACount,
                                                Sales,
                                                SalesCount,
                                                Discount,
                                                Surcharge,
                                                TranCount,
                                                Cash,
                                                CashCount,
                                                Card,
                                                CardCount,
                                                [GC],
                                                GCCount,
                                                IncentiveCard,
                                                IncentiveCardCount,
                                                CreditMemo,
                                                CreditMemoCount,
                                                CM_CashRefund,
                                                CM_CashRefundCount,
                                                ATD,
                                                ATDCount,
                                                VPlus,
                                                VPlusCount,
                                                [Misc],
                                                MiscCount,
                                                SN,
                                                PermitNo,
                                                M_I_N,
                                                Trans,
                                                Locked,
                                                VPlusCodeCount,
                                                Header1,
                                                Header2,
                                                Header3,
                                                TIN,
                                                ForOfflineMode,
                                                CapableOffline,
                                                WithEJournal,
                                                BankCommission,
                                                SupplierName,
                                                SupplierAddress1,
                                                SupplierAddress2,
                                                SupplierTIN,
                                                SupplierAccreditationNo,
                                                SupplierDateIssued,
                                                SupplierValidUntil,
                                                IsNewRegistered,
                                                IsNew,
                                                IsDisabled,
                                                HomeCredit ,
                                                HomeCreditCount,
                                                QRPay,
                                                QRPayCount)
                                    VALUES (
                                        '{fSqlFormat(rs.Fields("Counter").Value)}',
                                        {fNum(rs.Fields("TransactionCount").Value)},
                                        {fNum(rs.Fields("GrandTotal").Value)},
                                        {fNum(rs.Fields("ZZCount").Value)},
                                        '{fSqlFormat(rs.Fields("ResetCnt").Value)}',
                                        {fNum(rs.Fields("ResetTrans").Value)},
                                        '{fSqlFormat(rs.Fields("InvoiceNumberOld").Value)}',
                                         {fNum(rs.Fields("InvoiceNumberCnt").Value)},
                                        '{fSqlFormat(rs.Fields("InvoiceNumber").Value)}',
                                         {fNum(rs.Fields("RA").Value)},
                                         {fNum(rs.Fields("RACount").Value)},
                                         {fNum(rs.Fields("Sales").Value)},
                                         {fNum(rs.Fields("SalesCount").Value)},
                                         {fNum(rs.Fields("Discount").Value)},
                                         {fNum(rs.Fields("Surcharge").Value)},
                                         {fNum(rs.Fields("TranCount").Value)},
                                         {fNum(rs.Fields("Cash").Value)},
                                         {fNum(rs.Fields("CashCount").Value)},
                                         {fNum(rs.Fields("Card").Value)},
                                         {fNum(rs.Fields("CardCount").Value)},
                                         {fNum(rs.Fields("GC").Value)},
                                         {fNum(rs.Fields("GCCount").Value)},
                                         {fNum(rs.Fields("IncentiveCard").Value)},
                                         {fNum(rs.Fields("IncentiveCardCount").Value)},
                                         {fNum(rs.Fields("CreditMemo").Value)},
                                         {fNum(rs.Fields("CreditMemoCount").Value)},
                                         {fNum(rs.Fields("CM_CashRefund").Value)},
                                         {fNum(rs.Fields("CM_CashRefundCount").Value)},
                                         {fNum(rs.Fields("ATD").Value)},
                                         {fNum(rs.Fields("ATDCount").Value)},
                                         {fNum(rs.Fields("VPlus").Value)},
                                         {fNum(rs.Fields("VPlusCount").Value)},
                                         {fNum(rs.Fields("Misc").Value)},
                                         {fNum(rs.Fields("MiscCount").Value)},
                                        '{fSqlFormat(rs.Fields("SN").Value)}',
                                        '{fSqlFormat(rs.Fields("PermitNo").Value)}',
                                        '{fSqlFormat(rs.Fields("M_I_N").Value)}',
                                        {fNum(rs.Fields("Trans").Value)},
                                        {fNum(rs.Fields("Locked").Value)},
                                        {fNum(rs.Fields("VPlusCodeCount").Value)},
                                        '{fSqlFormat(rs.Fields("Header1").Value)}',
                                        '{fSqlFormat(rs.Fields("Header2").Value)}',
                                        '{fSqlFormat(rs.Fields("Header3").Value)}',
                                        '{fSqlFormat(rs.Fields("TIN").Value)}',
                                        {fNum(rs.Fields("ForOfflineMode").Value)},
                                        {fNum(rs.Fields("CapableOffline").Value)},
                                        {fNum(rs.Fields("WithEJournal").Value)},
                                        {fNum(rs.Fields("BankCommission").Value)},
                                        '{fSqlFormat(rs.Fields("SupplierName").Value)}',
                                        '{fSqlFormat(rs.Fields("SupplierAddress1").Value)}',
                                        '{fSqlFormat(rs.Fields("SupplierAddress2").Value)}',
                                        '{fSqlFormat(rs.Fields("SupplierTIN").Value)}',
                                        '{fSqlFormat(rs.Fields("SupplierAccreditationNo").Value)}',
                                        '{fSqlFormat(rs.Fields("SupplierDateIssued").Value)}',
                                        '{fSqlFormat(rs.Fields("SupplierValidUntil").Value)}',
                                        {fNum(rs.Fields("IsNewRegistered").Value)},
                                        {fNum(rs.Fields("IsNew").Value)},
                                        {fNum(rs.Fields("IsDisabled").Value)},
                                        {fNum(rs.Fields("HomeCredit").Value)},
                                        {fNum(rs.Fields("HomeCreditCount").Value)},
                                        {fNum(rs.Fields("QRPay").Value)},
                                        {fNum(rs.Fields("QRPayCount").Value)}
                                   );"

                    ConnServer.Execute(strSQL)
                Else

                    Dim strSQL As String = $"
                                    UPDATE tbl_PS_GT SET 
                                        TransactionCount = {fNum(rs.Fields("TransactionCount").Value)},
                                        GrandTotal = {fNum(rs.Fields("GrandTotal").Value)},
                                        ZZCount = {fNum(rs.Fields("ZZCount").Value)},
                                        ResetCnt = '{fSqlFormat(rs.Fields("ResetCnt").Value)}',
                                        ResetTrans = {fNum(rs.Fields("ResetTrans").Value)},
                                        InvoiceNumberOld = '{fSqlFormat(rs.Fields("InvoiceNumberOld").Value)}',
                                        InvoiceNumberCnt = {fNum(rs.Fields("InvoiceNumberCnt").Value)},
                                        InvoiceNumber = '{fSqlFormat(rs.Fields("InvoiceNumber").Value)}',
                                        RA = {fNum(rs.Fields("RA").Value)},
                                        RACount = {fNum(rs.Fields("RACount").Value)},
                                        Sales = {fNum(rs.Fields("Sales").Value)},
                                        SalesCount = {fNum(rs.Fields("SalesCount").Value)},
                                        Discount = {fNum(rs.Fields("Discount").Value)},
                                        Surcharge = {fNum(rs.Fields("Surcharge").Value)},
                                        TranCount = {fNum(rs.Fields("TranCount").Value)},
                                        Cash = {fNum(rs.Fields("Cash").Value)},
                                        CashCount = {fNum(rs.Fields("CashCount").Value)},
                                        Card = {fNum(rs.Fields("Card").Value)},
                                        CardCount = {fNum(rs.Fields("CardCount").Value)},
                                        [GC] = {fNum(rs.Fields("GC").Value)},
                                        GCCount = {fNum(rs.Fields("GCCount").Value)},
                                        IncentiveCard = {fNum(rs.Fields("IncentiveCard").Value)},
                                        IncentiveCardCount = {fNum(rs.Fields("IncentiveCardCount").Value)},
                                        CreditMemo = {fNum(rs.Fields("CreditMemo").Value)},
                                        CreditMemoCount = {fNum(rs.Fields("CreditMemoCount").Value)},
                                        CM_CashRefund = {fNum(rs.Fields("CM_CashRefund").Value)},
                                        CM_CashRefundCount = {fNum(rs.Fields("CM_CashRefundCount").Value)},
                                        ATD = {fNum(rs.Fields("ATD").Value)},
                                        ATDCount = {fNum(rs.Fields("ATDCount").Value)},
                                        VPlus = {fNum(rs.Fields("VPlus").Value)},
                                        VPlusCount = {fNum(rs.Fields("VPlusCount").Value)},
                                        [Misc] = {fNum(rs.Fields("Misc").Value)},
                                        MiscCount = {fNum(rs.Fields("MiscCount").Value)},
                                        SN = '{fSqlFormat(rs.Fields("SN").Value)}',
                                        PermitNo = '{fSqlFormat(rs.Fields("PermitNo").Value)}',
                                        M_I_N = '{fSqlFormat(rs.Fields("M_I_N").Value)}',
                                        Trans = {fNum(rs.Fields("Trans").Value)},
                                        Locked = {fNum(rs.Fields("Locked").Value)},
                                        VPlusCodeCount = {fNum(rs.Fields("VPlusCodeCount").Value)},
                                        Header1 = '{fSqlFormat(rs.Fields("Header1").Value)}',
                                        Header2 = '{fSqlFormat(rs.Fields("Header2").Value)}',
                                        Header3 = '{fSqlFormat(rs.Fields("Header3").Value)}',
                                        TIN = '{fSqlFormat(rs.Fields("TIN").Value)}',
                                        ForOfflineMode = {fNum(rs.Fields("ForOfflineMode").Value)},
                                        CapableOffline = {fNum(rs.Fields("CapableOffline").Value)},
                                        WithEJournal = {fNum(rs.Fields("WithEJournal").Value)},
                                        BankCommission = {fNum(rs.Fields("BankCommission").Value)},
                                        SupplierName = '{fSqlFormat(rs.Fields("SupplierName").Value)}',
                                        SupplierAddress1 = '{fSqlFormat(rs.Fields("SupplierAddress1").Value)}',
                                        SupplierAddress2 = '{fSqlFormat(rs.Fields("SupplierAddress2").Value)}',
                                        SupplierTIN = '{fSqlFormat(rs.Fields("SupplierTIN").Value)}',
                                        SupplierAccreditationNo = '{fSqlFormat(rs.Fields("SupplierAccreditationNo").Value)}',
                                        SupplierDateIssued = '{fSqlFormat(rs.Fields("SupplierDateIssued").Value)}',
                                        SupplierValidUntil = '{fSqlFormat(rs.Fields("SupplierValidUntil").Value)}',
                                        IsNewRegistered = {fNum(rs.Fields("IsNewRegistered").Value)},
                                        IsNew = {fNum(rs.Fields("IsNew").Value)},
                                        IsDisabled = {fNum(rs.Fields("IsDisabled").Value)},
                                        HomeCredit = {fNum(rs.Fields("HomeCredit").Value)},
                                        HomeCreditCount = {fNum(rs.Fields("HomeCreditCount").Value)},
                                        QRPay = {fNum(rs.Fields("QRPay").Value)},
                                        QRPayCount = {fNum(rs.Fields("QRPayCount").Value)}
                                        WHERE [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}';"

                    ConnServer.Execute(strSQL)

                End If

                rs.MoveNext()
            End While
        End If
    End Sub
    Public Sub Insert_tbl_PS_GT_ZZ(pb As ProgressBar, l As Label)


        rs = New ADODB.Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open($"SELECT * from tbl_PS_GT_ZZ ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_GT_ZZ :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim rx As New Recordset
                rx.Open($"SELECT TOP 1 * FROM tbl_PS_GT_ZZ WHERE [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}'", ConnServer, CursorTypeEnum.adOpenStatic)
                If (rx.RecordCount = 0) Then
                    Dim strSQL As String = $"INSERT INTO tbl_PS_GT_ZZ 
                                            ([Counter],
                                            [PSDate],
                                            ZZCount)
                                    VALUES ('{fSqlFormat(rs.Fields("Counter").Value)}',
                                         {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                        {fNum(rs.Fields("ZZCount").Value)}
                                   );"
                    ConnServer.Execute(strSQL)
                Else
                    Dim strSQL As String = $"UPDATE tbl_PS_GT_ZZ SET
                    [PSDate] = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                    ZZCount = {fNum(rs.Fields("ZZCount").Value)}
                    WHERE  [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}';"
                    ConnServer.Execute(strSQL)
                End If
                rs.MoveNext()
            End While
        End If
    End Sub
    Public Sub Insert_tbl_PS_E_Journal(pb As ProgressBar, l As Label)

        rs = New ADODB.Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open($"select * from tbl_PS_E_Journal where [Counter] = '{gbl_Counter}'  ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        Dim n As Integer = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_E_Journal  :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
                Dim rx As New Recordset
                rx.Open($"SELECT TOP 1 * FROM tbl_PS_E_Journal WHERE 
                                            PSNumber = '{fSqlFormat(rs.Fields("PSNumber").Value)}' and 
                                            PSDate={fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and 
                                            [Counter]='{fSqlFormat(rs.Fields("Counter").Value)}' and 
                                            Cashier='{fSqlFormat(rs.Fields("Cashier").Value)}'  ", ConnServer, CursorTypeEnum.adOpenStatic)

                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_PS_E_Journal  
                                            (PSNumber,
                                            PSDate,
                                            Cashier,
                                            [Counter],
                                            Series,
                                            ExactDate,
                                            Amount,
                                            SRem,
                                            TotalQty,
                                            TotalSales,
                                            TotalDiscount,
                                            TotalGC,
                                            TotalCard,
                                            TotalVPlus,
                                            TotalATD,
                                            Location,
                                            InvoiceNumber,
                                            VatPercent,
                                            VatSale,
                                            Vat,
                                            POSTableKey,
                                            TotalIncentiveCard,
                                            IsZeroRated,
                                            TotalCreditMemo,
                                            TotalHomeCredit,
                                            TotalQRPay)
                                    VALUES ('{fSqlFormat(rs.Fields("PSNumber").Value)}',
                                         {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                         '{fSqlFormat(rs.Fields("Cashier").Value)}',
                                         '{fSqlFormat(rs.Fields("Counter").Value)}',
                                         '{fSqlFormat(rs.Fields("Series").Value)}',
                                         {fDateIsEmpty(rs.Fields("ExactDate").Value.ToString())},
                                         {fNum(rs.Fields("Amount").Value)},
                                         '{fSqlFormat(rs.Fields("SRem").Value)}',
                                         {fNum(rs.Fields("TotalQty").Value)},
                                         {fNum(rs.Fields("TotalSales").Value)},
                                         {fNum(rs.Fields("TotalDiscount").Value)},
                                         {fNum(rs.Fields("TotalGC").Value)},
                                         {fNum(rs.Fields("TotalCard").Value)},
                                         {fNum(rs.Fields("TotalVPlus").Value)},
                                         {fNum(rs.Fields("TotalATD").Value)},
                                         '{fSqlFormat(rs.Fields("Location").Value)}',
                                         '{fSqlFormat(rs.Fields("InvoiceNumber").Value)}',
                                         '{fSqlFormat(rs.Fields("VatPercent").Value)}',
                                         {fNum(rs.Fields("VatSale").Value)},
                                         {fNum(rs.Fields("Vat").Value)},
                                         {fNum(rs.Fields("POSTableKey").Value)},
                                         {fNum(rs.Fields("TotalIncentiveCard").Value)},
                                         {fNum(rs.Fields("IsZeroRated").Value)},
                                         {fNum(rs.Fields("TotalCreditMemo").Value)},
                                         {fNum(rs.Fields("TotalHomeCredit").Value)},
                                         {fNum(rs.Fields("TotalQRPay").Value)}
                                   );"

                    ConnServer.Execute(strSQL)
                End If
                rs.MoveNext()
            End While
        End If
    End Sub
    Public Sub Insert_tbl_PS_E_Journal_Detail(pb As ProgressBar, l As Label)

        rs = New ADODB.Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open($"SELECT * from tbl_PS_E_Journal_Detail  WHERE [Counter] ='{gbl_Counter}' ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        Dim n As Integer = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_E_Journal_Detail :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
                Dim rx As New Recordset
                rx.Open($"SELECT TOP 1 * from tbl_PS_E_Journal_Detail 
                                            WHERE TransactionNumber={fSqlFormat(rs.Fields("TransactionNumber").Value)} and 
                                                    [Counter] ='{gbl_Counter}' and
                                                    PSDate= {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and 
                                                    Cashier = '{fSqlFormat(rs.Fields("Cashier").Value)}' and 
                                                    ItemCode='{fSqlFormat(rs.Fields("ItemCode").Value)}'", ConnServer, CursorTypeEnum.adOpenStatic)

                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_PS_E_Journal_Detail 
                                            (TransactionNumber,
                                            PSDate,
                                            [Counter],
                                            Cashier,
                                            ItemCode,
                                            ItemDescription,
                                            Quantity,
                                            GrossSRP,
                                            Discount,
                                            Surcharge,
                                            TotalGross,
                                            TotalDiscount,
                                            TotalSurcharge,
                                            TotalNet,
                                            Location,
                                            POSTableKey)
                                            VALUES (
                                                    '{fSqlFormat(rs.Fields("TransactionNumber").Value)}',
                                                    {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                                   '{fSqlFormat(rs.Fields("Counter").Value)}',
                                                   '{fSqlFormat(rs.Fields("Cashier").Value)}',
                                                   '{fSqlFormat(rs.Fields("ItemCode").Value)}',
                                                   '{fSqlFormat(rs.Fields("ItemDescription").Value)}',
                                                    {fNum(rs.Fields("Quantity").Value)},
                                                    {fNum(rs.Fields("GrossSRP").Value)},
                                                    {fNum(rs.Fields("Discount").Value)},
                                                    {fNum(rs.Fields("Surcharge").Value)},
                                                    {fNum(rs.Fields("TotalGross").Value)},
                                                    {fNum(rs.Fields("TotalDiscount").Value)},
                                                    {fNum(rs.Fields("TotalSurcharge").Value)},
                                                    {fNum(rs.Fields("TotalNet").Value)},
                                                   '{fSqlFormat(rs.Fields("Location").Value)}',
                                                    {fNum(rs.Fields("POSTableKey").Value)}
                                               );"

                    ConnServer.Execute(strSQL)

                Else
                    Dim strSQL As String = $"
                    UPDATE tbl_PS_E_Journal_Detail SET
                        TransactionNumber = '{fSqlFormat(rs.Fields("TransactionNumber").Value)}',
                        PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                        [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}',
                        Cashier = '{fSqlFormat(rs.Fields("Cashier").Value)}',
                        ItemCode = '{fSqlFormat(rs.Fields("ItemCode").Value)}',
                        ItemDescription = '{fSqlFormat(rs.Fields("ItemDescription").Value)}',
                        Quantity = {fNum(rs.Fields("Quantity").Value)},
                        GrossSRP = {fNum(rs.Fields("GrossSRP").Value)},
                        Discount = {fNum(rs.Fields("Discount").Value)},
                        Surcharge = {fNum(rs.Fields("Surcharge").Value)},
                        TotalGross = {fNum(rs.Fields("TotalGross").Value)},
                        TotalDiscount = {fNum(rs.Fields("TotalDiscount").Value)},
                        TotalSurcharge = {fNum(rs.Fields("TotalSurcharge").Value)},
                        TotalNet = {fNum(rs.Fields("TotalNet").Value)},
                        Location = '{fSqlFormat(rs.Fields("Location").Value)}',
                        POSTableKey = {fNum(rs.Fields("POSTableKey").Value)}
                        WHERE  TransactionNumber={fSqlFormat(rs.Fields("TransactionNumber").Value)} and 
                                [Counter] ='{gbl_Counter}' and
                                PSDate= {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and 
                                Cashier = '{fSqlFormat(rs.Fields("Cashier").Value)}' and 
                                ItemCode='{fSqlFormat(rs.Fields("ItemCode").Value)}';"

                    ConnServer.Execute(strSQL)
                End If

                rs.MoveNext()
            End While

        End If

    End Sub

    Public Sub Insert_tbl_PS_GT_Adjustment_EJournal(pb As ProgressBar, l As Label)

        rs = New ADODB.Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open($"select * from tbl_PS_GT_Adjustment_EJournal  where [Counter] = '{gbl_Counter}'  ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        Dim n As Integer = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_GT_Adjustment_EJournal :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
                Dim rx As New Recordset
                rx.Open($"select TOP 1 * from tbl_PS_GT_Adjustment_EJournal WHERE 
                                            PSNumber = '{fSqlFormat(rs.Fields("PSNumber").Value)}' and 
                                            PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and
                                            [Counter] = '{gbl_Counter}' ", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_PS_GT_Adjustment_EJournal 
                                            (PSNumber,
                                            PSDate,
                                            Cashier,
                                            [Counter],
                                            Series,
                                            ExactDate,
                                            Amount,
                                            SRem,
                                            TotalQty,
                                            TotalSales,
                                            TotalCash,
                                            TotalCard,
                                            TotalDiscount,
                                            TotalGC,
                                            TotalVPlus,
                                            TotalATD,
                                            Location,
                                            InvoiceNo,
                                            VatPercent,
                                            VatSale,
                                            Vat,
                                            POSTableKey,
                                            TotalIncentiveCard,
                                            IsZeroRated,
                                            UpdatedBy,
                                            LastUpdated,
                                            TotalCreditMemo)
                                            VALUES ('{fSqlFormat(rs.Fields("PSNumber").Value)}',
                                                   {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                                   '{fSqlFormat(rs.Fields("Cashier").Value)}',
                                                   '{fSqlFormat(rs.Fields("Counter").Value)}',                                      
                                                   '{fSqlFormat(rs.Fields("Series").Value)}',
                                                    {fDateIsEmpty(rs.Fields("ExactDate").Value.ToString())},
                                                    {fNum(rs.Fields("Amount").Value)},
                                                   '{fSqlFormat(rs.Fields("SRem").Value)}',
                                                    {fNum(rs.Fields("TotalQty").Value)},
                                                    {fNum(rs.Fields("TotalSales").Value)},
                                                    {fNum(rs.Fields("TotalCash").Value)},
                                                    {fNum(rs.Fields("TotalCard").Value)},
                                                    {fNum(rs.Fields("TotalDiscount").Value)},
                                                    {fNum(rs.Fields("TotalGC").Value)},
                                                    {fNum(rs.Fields("TotalVPlus").Value)},
                                                    {fNum(rs.Fields("TotalATD").Value)},
                                                   '{fSqlFormat(rs.Fields("Location").Value)}',
                                                   '{fSqlFormat(rs.Fields("InvoiceNo").Value)}',
                                                   '{fSqlFormat(rs.Fields("VatPercent").Value)}',
                                                    {fNum(rs.Fields("VatSale").Value)},
                                                    {fNum(rs.Fields("Vat").Value)},
                                                    {fNum(rs.Fields("POSTableKey").Value)},
                                                    {fNum(rs.Fields("TotalIncentiveCard").Value)},
                                                    {fNum(rs.Fields("IsZeroRated").Value)},
                                                    '{fSqlFormat(rs.Fields("UpdatedBy").Value)}',
                                                    {fDateIsEmpty(rs.Fields("LastUpdated").Value.ToString())},
                                                    {fNum(rs.Fields("TotalCreditMemo").Value)}
                                               );"

                    ConnServer.Execute(strSQL)

                Else
                    Dim strSQL As String = $"UPDATE tbl_PS_GT_Adjustment_EJournal SET
                                PSNumber = '{fSqlFormat(rs.Fields("PSNumber").Value)}',
                                PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                Cashier = '{fSqlFormat(rs.Fields("Cashier").Value)}',
                                [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}',
                                Series = '{fSqlFormat(rs.Fields("Series").Value)}',
                                ExactDate = {fDateIsEmpty(rs.Fields("ExactDate").Value.ToString())},
                                Amount = {fNum(rs.Fields("Amount").Value)},
                                SRem = '{fSqlFormat(rs.Fields("SRem").Value)}',
                                TotalQty = {fNum(rs.Fields("TotalQty").Value)},
                                TotalSales = {fNum(rs.Fields("TotalSales").Value)},
                                TotalCash = {fNum(rs.Fields("TotalCash").Value)},
                                TotalCard = {fNum(rs.Fields("TotalCard").Value)},
                                TotalDiscount = {fNum(rs.Fields("TotalDiscount").Value)},
                                TotalGC = {fNum(rs.Fields("TotalGC").Value)},
                                TotalVPlus = {fNum(rs.Fields("TotalVPlus").Value)},
                                TotalATD = {fNum(rs.Fields("TotalATD").Value)},
                                Location = '{fSqlFormat(rs.Fields("Location").Value)}',
                                InvoiceNo = '{fSqlFormat(rs.Fields("InvoiceNo").Value)}',
                                VatPercent = '{fSqlFormat(rs.Fields("VatPercent").Value)}',
                                VatSale = {fNum(rs.Fields("VatSale").Value)},
                                Vat = {fNum(rs.Fields("Vat").Value)},
                                POSTableKey = {fNum(rs.Fields("POSTableKey").Value)},
                                TotalIncentiveCard = {fNum(rs.Fields("TotalIncentiveCard").Value)},
                                IsZeroRated = {fNum(rs.Fields("IsZeroRated").Value)},
                                UpdatedBy = '{fSqlFormat(rs.Fields("UpdatedBy").Value)}',
                                LastUpdated = {fDateIsEmpty(rs.Fields("LastUpdated").Value.ToString())},
                                TotalCreditMemo = {fNum(rs.Fields("TotalCreditMemo").Value)}
                                WHERE PSNumber = '{fSqlFormat(rs.Fields("PSNumber").Value)}' and PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and [Counter] = '{gbl_Counter}';"

                    ConnServer.Execute(strSQL)
                End If
                rs.MoveNext()
            End While

        End If

    End Sub
    Public Sub Insert_tbl_PS_GT_Adjustment_EJournal_Detail(pb As ProgressBar, l As Label)

        rs = New Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open($"SELECT * from tbl_PS_GT_Adjustment_EJournal_Detail WHERE  [Counter] = '{gbl_Counter}'  ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        Dim n As Integer = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_GT_Adjustment_EJournal_Detail :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If

                Dim rx As New Recordset
                rx.Open($"SELECT TOP 1 * from tbl_PS_GT_Adjustment_EJournal_Detail 
                                                WHERE TransactionNumber = '{fSqlFormat(rs.Fields("TransactionNumber").Value)}' and 
                                                PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and 
                                                [Counter] = '{gbl_Counter}' and 
                                                Cashier = '{fSqlFormat(rs.Fields("Cashier").Value)}' and 
                                                ItemCode = '{fSqlFormat(rs.Fields("ItemCode").Value)}' ", ConnServer, CursorTypeEnum.adOpenStatic)

                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_PS_GT_Adjustment_EJournal_Detail 
                                            (
                                            TransactionNumber,
                                            PSDate,
                                            [Counter],
                                            Cashier,                        
                                            ItemCode,
                                            ItemDescription,
                                            Quantity,
                                            GrossSRP,
                                            Discount,
                                            Surcharge,
                                            TotalGross,
                                            TotalDiscount,
                                            TotalSurcharge,
                                            TotalNet,
                                            Location,
                                            POSTableKey)
                                            VALUES (
                                                   '{fSqlFormat(rs.Fields("TransactionNumber").Value)}',
                                                    {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                                   '{fSqlFormat(rs.Fields("Counter").Value)}',
                                                   '{fSqlFormat(rs.Fields("Cashier").Value)}',                                      
                                                   '{fSqlFormat(rs.Fields("ItemCode").Value)}',
                                                   '{fSqlFormat(rs.Fields("ItemDescription").Value)}',
                                                    {fNum(rs.Fields("Quantity").Value)},
                                                    {fNum(rs.Fields("GrossSRP").Value)},
                                                    {fNum(rs.Fields("Discount").Value)},
                                                    {fNum(rs.Fields("Surcharge").Value)},
                                                    {fNum(rs.Fields("TotalGross").Value)},
                                                    {fNum(rs.Fields("TotalDiscount").Value)},
                                                    {fNum(rs.Fields("TotalSurcharge").Value)},
                                                    {fNum(rs.Fields("TotalNet").Value)},
                                                   '{fSqlFormat(rs.Fields("Location").Value)}',
                                                    {fNum(rs.Fields("POSTableKey").Value)}
                                               );"

                    ConnServer.Execute(strSQL)
                Else
                    Dim strSQL As String = $"
                        UPDATE tbl_PS_GT_Adjustment_EJournal_Detail SET
                        TransactionNumber = '{fSqlFormat(rs.Fields("TransactionNumber").Value)}',
                        PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                        [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}',
                        Cashier = '{fSqlFormat(rs.Fields("Cashier").Value)}',
                        ItemCode = '{fSqlFormat(rs.Fields("ItemCode").Value)}',
                        ItemDescription = '{fSqlFormat(rs.Fields("ItemDescription").Value)}',
                        Quantity = {fNum(rs.Fields("Quantity").Value)},
                        GrossSRP = {fNum(rs.Fields("GrossSRP").Value)},
                        Discount = {fNum(rs.Fields("Discount").Value)},
                        Surcharge = {fNum(rs.Fields("Surcharge").Value)},
                        TotalGross = {fNum(rs.Fields("TotalGross").Value)},
                        TotalDiscount = {fNum(rs.Fields("TotalDiscount").Value)},
                        TotalSurcharge = {fNum(rs.Fields("TotalSurcharge").Value)},
                        TotalNet = {fNum(rs.Fields("TotalNet").Value)},
                        Location = '{fSqlFormat(rs.Fields("Location").Value)}',
                        POSTableKey = {fNum(rs.Fields("POSTableKey").Value)}
                        WHERE 
                        TransactionNumber = '{fSqlFormat(rs.Fields("TransactionNumber").Value)}' and 
                        PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and 
                        [Counter] = '{gbl_Counter}' and
                        Cashier = '{fSqlFormat(rs.Fields("Cashier").Value)}' and
                        ItemCode = '{fSqlFormat(rs.Fields("ItemCode").Value)}';"

                    ConnServer.Execute(strSQL)

                End If

                rs.MoveNext()
            End While

        End If

    End Sub

    Public Sub Insert_tbl_PaidOutDenominations(pb As ProgressBar, l As Label)
        ConnServer.Execute("SET IDENTITY_INSERT tbl_PaidOutDenominations ON;")
        rs = New ADODB.Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open($"select * from tbl_PaidOutDenominations ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PaidOutDenominations :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()

                Dim rx As New Recordset
                rx.Open($"select TOP 1 * from tbl_PaidOutDenominations  where [DenomPK] = {fNum(rs.Fields("DenomPK").Value)} ", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_PaidOutDenominations 
                                            ([DenomPK],
                                            [Denominations_Code],
                                            Denominations,
                                            [Type],
                                            [Active]  )
                                    VALUES (
                                        {fNum(rs.Fields("DenomPK").Value)},
                                        '{fSqlFormat(rs.Fields("Denominations_Code").Value)}',
                                        {fNum(rs.Fields("Denominations").Value)},
                                        {fNum(rs.Fields("Type").Value)},
                                        {fNum(rs.Fields("Active").Value)}
                                   );"

                    ConnServer.Execute(strSQL)
                End If


                rs.MoveNext()
            End While

        End If
        ConnServer.Execute("SET IDENTITY_INSERT tbl_PaidOutDenominations OFF;")
    End Sub

    Public Sub Insert_tbl_PaidOutTransactions(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 1
        Dim n As Integer = 0
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PaidOutTransactions WHERE MachineNo = '{gbl_Counter}'", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PaidOutTransactions  :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If

                Dim rx As New Recordset
                rx.Open($"SELECT * FROM tbl_PaidOutTransactions WHERE TransDate =  {fDateIsEmpty(rs.Fields("TransDate").Value.ToString())} and TransTime = '{fSqlFormat(rs.Fields("TransTime").Value)}'  and MachineNo = '{fSqlFormat(rs.Fields("MachineNo").Value)}' ", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_PaidOutTransactions 
                                                (   PaidOutPK,
                                                    TransDate,
                                                    TransTime,
                                                    CtrlNo,
                                                    OOrder,
                                                    CashierCode,
                                                    CashierName,
                                                    CollectorCode,
                                                    CollectorName,
                                                    MachineNo,
                                                    Total,
                                                    YYear,
                                                    Series,
                                                    IsPosted,
                                                    IsChecked,
                                                    Total_Previous,
                                                    SessionPK,
                                                    IsUsed)
                                                VALUES ({fNum(rs.Fields("PaidOutPK").Value)},      
                                                         {fDateIsEmpty(rs.Fields("TransDate").Value.ToString())},
                                                        '{fSqlFormat(rs.Fields("TransTime").Value)}',
                                                        '{fSqlFormat(rs.Fields("CtrlNo").Value)}',
                                                         {fNum(rs.Fields("OOrder").Value)},                                                     '{fSqlFormat(rs.Fields("Track1").Value)}',
                                                        '{fSqlFormat(rs.Fields("CashierCode").Value)}',
                                                        '{fSqlFormat(rs.Fields("Type").Value)}',
                                                        '{fSqlFormat(rs.Fields("CashierName").Value)}',
                                                        '{fSqlFormat(rs.Fields("CollectorCode").Value)}',
                                                        '{fSqlFormat(rs.Fields("CollectorName").Value)}',      
                                                        '{fSqlFormat(rs.Fields("MachineNo").Value)}',     
                                                         {fNum(rs.Fields("Total").Value)},
                                                         {fNum(rs.Fields("YYear").Value)},                                  
                                                         {fNum(rs.Fields("Series").Value)},
                                                         {fNum(rs.Fields("IsPosted").Value)},
                                                         {fNum(rs.Fields("IsChecked").Value)},
                                                         {fNum(rs.Fields("Total_Previous").Value)},
                                                         {fNum(rs.Fields("SessionPK").Value)},
                                                         {fNum(rs.Fields("IsUsed").Value)}
                                                );"

                    ConnServer.Execute(strSQL)
                End If

                rs.MoveNext()
            End While
        End If

    End Sub

End Module
