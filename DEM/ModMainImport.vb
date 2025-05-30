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
                rx.Open($"select TOP 1 * from tbl_Items as i WHere i.PK = {rs.Fields("PK").Value}", ConnServer, CursorTypeEnum.adOpenStatic)
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



End Module
