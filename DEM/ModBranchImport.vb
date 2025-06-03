Imports ADODB
Module ModBranchImport


    Public Sub Branch_Insert_tbl_GiftCert_List(pb As ProgressBar, l As Label)

        rs = New Recordset
        rs.Open($"select * from tbl_GiftCert_List ", ConnLocal, CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_GiftCert_List :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()

                Dim rx As New Recordset
                rx.Open($"SELECT TOP 1 * FROM tbl_GiftCert_List WHERE ValidFrom = {fDateIsEmpty(rs.Fields("ValidFrom").Value.ToString())} and 
                                                                      ValidTo = {fDateIsEmpty(rs.Fields("ValidTo").Value.ToString())} And
                                                                      GCNumber = { fNum(rs.Fields("GCNumber").Value)} And
                                                                      DateAdded = {fDateIsEmpty(rs.Fields("DateAdded").Value.ToString())} ", ConnServer, CursorTypeEnum.adOpenStatic)

                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_GiftCert_List 
                                    (GCNumber,
                                    Amount,
                                    Customer,
                                    ValidFrom,
                                    ValidTo,
                                    DateAdded,
                                    Used,
                                    DateUsed)
                                    VALUES ({rs.Fields("GCNumber").Value},
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

    End Sub
    Public Sub Branch_Insert_tbl_VPlus_Codes(pb As ProgressBar, l As Label)

        Dim n As Integer = 0
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_VPlus_Codes ", ConnLocal, CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_VPlus_Codes :" & pb.Maximum & "/" & pb.Value

                n = 0
                    Application.DoEvents()

                Dim rx As New Recordset
                rx.Open($"select TOP 1 * from tbl_VPlus_Codes WHERE Codes ='{fSqlFormat(rs.Fields("Codes").Value)}'  ", ConnServer, CursorTypeEnum.adOpenStatic)
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
    Public Sub Branch_Insert_tbl_VPlus_Codes_Validity(pb As ProgressBar, l As Label)

        rs = New Recordset
        rs.Open($"select * from tbl_VPlus_Codes_Validity  ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        Dim n As Integer = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_VPlus_Codes_Validity :" & pb.Maximum & "/" & pb.Value

                Application.DoEvents()

                    Dim rx As New Recordset
                rx.Open($"select TOP 1 * FROM tbl_VPlus_Codes_Validity WHERE 
                        Codes='{fSqlFormat(rs.Fields("Codes").Value)}' and 
                        DateStarted={fDateIsEmpty(rs.Fields("DateStarted").Value.ToString())} and 
                        DateExpired={fDateIsEmpty(rs.Fields("DateExpired").Value.ToString())} ", ConnServer, CursorTypeEnum.adOpenStatic)

                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_VPlus_Codes_Validity 
                                    (Codes,
                                    DateStarted,
                                    DateExpired,
                                    GracePeriod)
                                    VALUES ('{fSqlFormat(rs.Fields("Codes").Value)}',  
                                    {fDateIsEmpty(rs.Fields("DateStarted").Value.ToString())},
                                    {fDateIsEmpty(rs.Fields("DateExpired").Value.ToString())},    
                                    {fDateIsEmpty(rs.Fields("GracePeriod").Value.ToString())});"

                    ConnServer.Execute(strSQL)
                End If
                rs.MoveNext()
            End While

        End If

    End Sub
    Public Sub Branch_Insert_tbl_PS_GT(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 5

        rs = New ADODB.Recordset
        rs.Open($"SELECT * FROM tbl_PS_GT ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
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
                If rx.RecordCount = 0 Then
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

                    Dim strSQL As String = $"UPDATE tbl_PS_GT SET 
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
    Public Sub Branch_Insert_tbl_PS_GT_ZZ(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 5

        rs = New Recordset
        rs.CursorLocation = CursorLocationEnum.adUseClient
        rs.Open($"select * from tbl_PS_GT_ZZ ", ConnLocal, CursorTypeEnum.adOpenStatic)
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
                                    VALUES (
                                        '{fSqlFormat(rs.Fields("Counter").Value)}',
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
    Public Sub Branch_Insert_tbl_PS_E_Journal(pb As ProgressBar, l As Label)

        rs = New ADODB.Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open($"select * from tbl_PS_E_Journal ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        Dim n As Integer = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_E_Journal  :" & pb.Maximum & "/" & pb.Value

                Application.DoEvents()

                Dim rx As New Recordset
                rx.Open($"SELECT TOP 1 * From tbl_PS_E_Journal WHERE 
                        PSNumber = '{fSqlFormat(rs.Fields("PSNumber").Value)}' and
                        Cashier= '{fSqlFormat(rs.Fields("Cashier").Value)}' and 
                        PSDate={fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and 
                        Counter='{fSqlFormat(rs.Fields("Counter").Value)}'", ConnServer, CursorTypeEnum.adOpenStatic)
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

                Else
                    Dim strSQL As String = $"UPDATE tbl_PS_E_Journal SET
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
                            TotalDiscount = {fNum(rs.Fields("TotalDiscount").Value)},
                            TotalGC = {fNum(rs.Fields("TotalGC").Value)},
                            TotalCard = {fNum(rs.Fields("TotalCard").Value)},
                            TotalVPlus = {fNum(rs.Fields("TotalVPlus").Value)},
                            TotalATD = {fNum(rs.Fields("TotalATD").Value)},
                            Location = '{fSqlFormat(rs.Fields("Location").Value)}',
                            InvoiceNumber = '{fSqlFormat(rs.Fields("InvoiceNumber").Value)}',
                            VatPercent = '{fSqlFormat(rs.Fields("VatPercent").Value)}',
                            VatSale = {fNum(rs.Fields("VatSale").Value)},
                            Vat = {fNum(rs.Fields("Vat").Value)},
                            POSTableKey = {fNum(rs.Fields("POSTableKey").Value)},
                            TotalIncentiveCard = {fNum(rs.Fields("TotalIncentiveCard").Value)},
                            IsZeroRated = {fNum(rs.Fields("IsZeroRated").Value)},
                            TotalCreditMemo = {fNum(rs.Fields("TotalCreditMemo").Value)},
                            TotalHomeCredit = {fNum(rs.Fields("TotalHomeCredit").Value)},
                            TotalQRPay = {fNum(rs.Fields("TotalQRPay").Value)}
                            WHERE PSNumber = '{fSqlFormat(rs.Fields("PSNumber").Value)}' and
                            Cashier= '{fSqlFormat(rs.Fields("Cashier").Value)}' and 
                            PSDate={fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and 
                            Counter='{fSqlFormat(rs.Fields("Counter").Value)}'"

                    ConnServer.Execute(strSQL)
                End If

                rs.MoveNext()
            End While
        End If
    End Sub
    Public Sub Branch_Insert_tbl_PS_E_Journal_Detail(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 1
        Dim toDate As String = Now.Date.ToShortDateString()
        Dim FromDate As String = Now.Date.AddYears(-1).ToShortDateString()
        rs = New ADODB.Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open($"select * from tbl_PS_E_Journal_Detail  ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        Dim n As Integer = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_E_Journal_Detail :" & pb.Maximum & "/" & pb.Value

                Application.DoEvents()


                    Dim rx As New Recordset
                rx.Open($"select TOP 1 * from tbl_PS_E_Journal_Detail WHERE 
                                                                        TransactionNumber='{fSqlFormat(rs.Fields("TransactionNumber").Value)}' and 
                                                                        PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and 
                                                                        [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}' and 
                                                                        Cashier='{fSqlFormat(rs.Fields("Cashier").Value)}' and 
                                                                        ItemCode= '{fSqlFormat(rs.Fields("ItemCode").Value)}' and 
                                                                        POSTableKey = {fNum(rs.Fields("POSTableKey").Value)} ", ConnServer, CursorTypeEnum.adOpenStatic)
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
                                            VALUES ('{fSqlFormat(rs.Fields("TransactionNumber").Value)}',
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
                                                    {fNum(rs.Fields("POSTableKey").Value)});"

                    ConnServer.Execute(strSQL)

                Else

                    Dim strSQL As String = $"UPDATE tbl_PS_E_Journal_Detail SET
                            TransactionNumber = '{fSqlFormat(rs.Fields("TransactionNumber").Value)}',
                            PSDate            = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                            [Counter]         = '{fSqlFormat(rs.Fields("Counter").Value)}',
                            Cashier           = '{fSqlFormat(rs.Fields("Cashier").Value)}',
                            ItemCode          = '{fSqlFormat(rs.Fields("ItemCode").Value)}',
                            ItemDescription   = '{fSqlFormat(rs.Fields("ItemDescription").Value)}',
                            Quantity          = {fNum(rs.Fields("Quantity").Value)},
                            GrossSRP          = {fNum(rs.Fields("GrossSRP").Value)},
                            Discount          = {fNum(rs.Fields("Discount").Value)},
                            Surcharge         = {fNum(rs.Fields("Surcharge").Value)},
                            TotalGross        = {fNum(rs.Fields("TotalGross").Value)},
                            TotalDiscount     = {fNum(rs.Fields("TotalDiscount").Value)},
                            TotalSurcharge    = {fNum(rs.Fields("TotalSurcharge").Value)},
                            TotalNet          = {fNum(rs.Fields("TotalNet").Value)},
                            Location          = '{fSqlFormat(rs.Fields("Location").Value)}',
                            POSTableKey       = {fNum(rs.Fields("POSTableKey").Value)}
                            WHERE TransactionNumber='{fSqlFormat(rs.Fields("TransactionNumber").Value)}' and 
                                PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and 
                                [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}' and 
                                Cashier='{fSqlFormat(rs.Fields("Cashier").Value)}' and 
                                ItemCode= '{fSqlFormat(rs.Fields("ItemCode").Value)}' and 
                                POSTableKey = {fNum(rs.Fields("POSTableKey").Value)} ;"

                    ConnServer.Execute(strSQL)
                End If

                rs.MoveNext()
            End While

        End If

    End Sub
    Public Sub Branch_Insert_tbl_PS_EmployeeATD(pb As ProgressBar, l As Label)

        Dim n As Integer = 0
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PS_EmployeeATD ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_EmployeeATD :" & pb.Maximum & "/" & pb.Value

                Application.DoEvents()

                    Dim rx As New Recordset
                rx.Open($"select * from tbl_PS_EmployeeATD Where TransactionNumber = '{fSqlFormat(rs.Fields("TransactionNumber").Value)}' and
                                                            PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and
                                                            [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}' and 
                                                            Cashier='{fSqlFormat(rs.Fields("Cashier").Value)}' and 
                                                            ATDNumber='{fSqlFormat(rs.Fields("ATDNumber").Value)}' ", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then

                    Dim strSQL As String = $"INSERT INTO tbl_PS_EmployeeATD 
                                                (TransactionNumber,
                                                PSDate,
                                                [Counter],
                                                Cashier,
                                                ATDNumber,
                                                EmpNo,
                                                Amount)
                                                VALUES (      
                                                '{fSqlFormat(rs.Fields("TransactionNumber").Value)}',
                                                 {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                                '{fSqlFormat(rs.Fields("Counter").Value)}',
                                                '{fSqlFormat(rs.Fields("Cashier").Value)}',
                                                '{fSqlFormat(rs.Fields("ATDNumber").Value)}',
                                                {fNum(rs.Fields("EmpNo").Value)},
                                                {fNum(rs.Fields("Amount").Value)});"

                    ConnServer.Execute(strSQL)
                Else
                    Dim strSQL As String = $"UPDATE tbl_PS_EmployeeATD SET
                            EmpNo             = {fNum(rs.Fields("EmpNo").Value)},
                            Amount            = {fNum(rs.Fields("Amount").Value)}
                            Where TransactionNumber = '{fSqlFormat(rs.Fields("TransactionNumber").Value)}' and
                            PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and
                            [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}' and 
                            Cashier='{fSqlFormat(rs.Fields("Cashier").Value)}' and 
                            ATDNumber='{fSqlFormat(rs.Fields("ATDNumber").Value)}'"

                    ConnServer.Execute(strSQL)
                End If

                rs.MoveNext()
            End While
        End If

    End Sub
    Public Sub Branch_Insert_tbl_GiftCert_Payment(pb As ProgressBar, l As Label)


        rs = New ADODB.Recordset
        rs.Open($"SELECT * from tbl_GiftCert_Payment  ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_GiftCert_Payment :" & pb.Maximum & "/" & pb.Value

                Application.DoEvents()
                Dim rx As New Recordset
                rx.Open($"SELECT * from tbl_GiftCert_Payment WHERE PSNumber='{fSqlFormat(rs.Fields("PSNumber").Value)}' and PSDate= {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and [Counter]='{fSqlFormat(rs.Fields("Counter").Value)}' and Cashier='{fSqlFormat(rs.Fields("Cashier").Value)}' and GCNumber='{fSqlFormat(rs.Fields("GCNumber").Value)}'", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_GiftCert_Payment 
                                                (PSNumber,
                                                PSDate,
                                                [Counter],
                                                Cashier,
                                                GCNumber,
                                                GCAmount,
                                                Posted)
                                                VALUES (     
                                                '{fSqlFormat(rs.Fields("PSNumber").Value)}',
                                                 {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                                '{fSqlFormat(rs.Fields("Counter").Value)}',
                                                '{fSqlFormat(rs.Fields("Cashier").Value)}',
                                                '{fSqlFormat(rs.Fields("GCNumber").Value)}',
                                                {fNum(rs.Fields("GCAmount").Value)},
                                                {fNum(rs.Fields("Posted").Value)});"

                    ConnServer.Execute(strSQL)
                Else
                    Dim strSQL As String = $"UPDATE tbl_GiftCert_Payment SET
                            GCAmount = {fNum(rs.Fields("GCAmount").Value)},
                            Posted   = {fNum(rs.Fields("Posted").Value)}
                            WHERE PSNumber='{fSqlFormat(rs.Fields("PSNumber").Value)}' and 
                            PSDate= {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and 
                            [Counter]='{fSqlFormat(rs.Fields("Counter").Value)}' and 
                            Cashier='{fSqlFormat(rs.Fields("Cashier").Value)}' and 
                            GCNumber='{fSqlFormat(rs.Fields("GCNumber").Value)}';"

                    ConnServer.Execute(strSQL)
                End If


                rs.MoveNext()
            End While
        End If

    End Sub
    Public Sub Branch_Insert_tbl_VPlus_Purchases_Points(pb As ProgressBar, l As Label)


        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_VPlus_Purchases_Points ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_VPlus_Purchases_Points :" & pb.Maximum & "/" & pb.Value

                Application.DoEvents()

                Dim rx As New Recordset
                rx.Open($"select TOP 1 * from tbl_VPlus_Purchases_Points WHERE TransactionNo = '{fSqlFormat(rs.Fields("TransactionNo").Value)}' and
                    VDate= {fDateIsEmpty(rs.Fields("VDate").Value.ToString())} and 
                    VPlusCodes = '{fSqlFormat(rs.Fields("VPlusCodes").Value)}' and 
                    [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}' and 
                    Cashier= '{fSqlFormat(rs.Fields("Cashier").Value)}'", ConnServer, CursorTypeEnum.adOpenStatic)

                If rx.RecordCount = 0 Then

                    Dim strSQL As String = $"INSERT INTO tbl_VPlus_Purchases_Points 
                                                (TransactionNo,
                                                VDate,
                                                VPlusCodes,
                                                [Counter],
                                                Cashier,
                                                Cash,
                                                Card,
                                                [GC],
                                                [ATD],
                                                [PointsPay],
                                                [Location],
                                                [Posted])
                                                VALUES ('{fSqlFormat(rs.Fields("TransactionNo").Value)}',
                                                 {fDateIsEmpty(rs.Fields("VDate").Value.ToString())},
                                                '{fSqlFormat(rs.Fields("VPlusCodes").Value)}',
                                                '{fSqlFormat(rs.Fields("Counter").Value)}',
                                                '{fSqlFormat(rs.Fields("Cashier").Value)}',
                                                {fNum(rs.Fields("Cash").Value)},
                                                {fNum(rs.Fields("Card").Value)},
                                                {fNum(rs.Fields("GC").Value)},
                                                {fNum(rs.Fields("ATD").Value)},
                                                {fNum(rs.Fields("PointsPay").Value)},
                                                '{fSqlFormat(rs.Fields("Location").Value)}',
                                                {fNum(rs.Fields("Posted").Value)}
                                                );"

                    ConnLocal.Execute(strSQL)
                Else
                    Dim strSQL As String = $"UPDATE tbl_VPlus_Purchases_Points SET
                            Cash           = {fNum(rs.Fields("Cash").Value)},
                            Card           = {fNum(rs.Fields("Card").Value)},
                            [GC]           = {fNum(rs.Fields("GC").Value)},
                            [ATD]          = {fNum(rs.Fields("ATD").Value)},
                            PointsPay      = {fNum(rs.Fields("PointsPay").Value)},
                            [Location]     = '{fSqlFormat(rs.Fields("Location").Value)}',
                            [Posted]       = {fNum(rs.Fields("Posted").Value)}
                            WHERE TransactionNo = '{fSqlFormat(rs.Fields("TransactionNo").Value)}' and VDate= {fDateIsEmpty(rs.Fields("VDate").Value.ToString())} and VPlusCodes = '{fSqlFormat(rs.Fields("VPlusCodes").Value)}' and [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}' and Cashier= '{fSqlFormat(rs.Fields("Cashier").Value)}';"
                    ConnServer.Execute(strSQL)
                End If



                rs.MoveNext()
            End While
        End If

    End Sub
    Public Sub Branch_Insert_tbl_PS_Tmp(pb As ProgressBar, l As Label)


        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PS_Tmp  ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_Tmp :" & pb.Maximum & "/" & pb.Value

                Application.DoEvents()
                Dim rx As New Recordset
                rx.Open($"select TOP 1 * from tbl_PS_Tmp WHERE PSNumber= '{fSqlFormat(rs.Fields("PSNumber").Value)}' and PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and Cashier= '{fSqlFormat(rs.Fields("Cashier").Value)}' and [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}' and [Series]= '{fSqlFormat(rs.Fields("Series").Value)}' ", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_PS_Tmp 
                                                (PSNumber,
                                                PSDate,
                                                PSDateActual,
                                                Cashier,
                                                [Counter],
                                                [Series],
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
                                                Posted,
                                                POSTableKey,
                                                TotalIncentiveCard,
                                                IsZeroRated,
                                                TotalCreditMemo,
                                                TotalHomeCredit,
                                                TotalQRPay)
                                                VALUES (      
                                                '{fSqlFormat(rs.Fields("PSNumber").Value)}',
                                                 {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                                 {fDateIsEmpty(rs.Fields("PSDateActual").Value.ToString())},
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
                                                {fNum(rs.Fields("Posted").Value)},
                                                {fNum(rs.Fields("POSTableKey").Value)},
                                                {fNum(rs.Fields("TotalIncentiveCard").Value)},
                                                {fNum(rs.Fields("IsZeroRated").Value)},
                                                {fNum(rs.Fields("TotalCreditMemo").Value)},
                                                {fNum(rs.Fields("TotalHomeCredit").Value)},
                                                {fNum(rs.Fields("TotalQRPay").Value)}
                                                );"

                    ConnServer.Execute(strSQL)

                Else
                    Dim strSQL As String = $"UPDATE tbl_PS_Tmp SET
                            PSNumber            = '{fSqlFormat(rs.Fields("PSNumber").Value)}',
                            PSDate              = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                            PSDateActual        = {fDateIsEmpty(rs.Fields("PSDateActual").Value.ToString())},
                            Cashier             = '{fSqlFormat(rs.Fields("Cashier").Value)}',
                            [Counter]           = '{fSqlFormat(rs.Fields("Counter").Value)}',
                            [Series]            = '{fSqlFormat(rs.Fields("Series").Value)}',
                            ExactDate           = {fDateIsEmpty(rs.Fields("ExactDate").Value.ToString())},
                            Amount              = {fNum(rs.Fields("Amount").Value)},
                            SRem                = '{fSqlFormat(rs.Fields("SRem").Value)}',
                            TotalQty            = {fNum(rs.Fields("TotalQty").Value)},
                            TotalSales          = {fNum(rs.Fields("TotalSales").Value)},
                            TotalDiscount       = {fNum(rs.Fields("TotalDiscount").Value)},
                            TotalGC             = {fNum(rs.Fields("TotalGC").Value)},
                            TotalCard           = {fNum(rs.Fields("TotalCard").Value)},
                            TotalVPlus          = {fNum(rs.Fields("TotalVPlus").Value)},
                            TotalATD            = {fNum(rs.Fields("TotalATD").Value)},
                            Location            = '{fSqlFormat(rs.Fields("Location").Value)}',
                            InvoiceNumber       = '{fSqlFormat(rs.Fields("InvoiceNumber").Value)}',
                            Posted              = {fNum(rs.Fields("Posted").Value)},
                            POSTableKey         = {fNum(rs.Fields("POSTableKey").Value)},
                            TotalIncentiveCard  = {fNum(rs.Fields("TotalIncentiveCard").Value)},
                            IsZeroRated         = {fNum(rs.Fields("IsZeroRated").Value)},
                            TotalCreditMemo     = {fNum(rs.Fields("TotalCreditMemo").Value)},
                            TotalHomeCredit     = {fNum(rs.Fields("TotalHomeCredit").Value)},
                            TotalQRPay          = {fNum(rs.Fields("TotalQRPay").Value)}
                            WHERE PSNumber= '{fSqlFormat(rs.Fields("PSNumber").Value)}' and PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and Cashier= '{fSqlFormat(rs.Fields("Cashier").Value)}' and [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}' and [Series]= '{fSqlFormat(rs.Fields("Series").Value)}';"

                    ConnServer.Execute(strSQL)

                End If



                rs.MoveNext()
            End While
        End If

    End Sub

    Public Sub Branch_Insert_tbl_PS_ItemsSold_Tmp(pb As ProgressBar, l As Label)

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PS_ItemsSold_Tmp ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_ItemsSold_Tmp :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()

                Dim rx As New Recordset
                rx.Open($"select * from tbl_PS_ItemsSold_Tmp WHERE  TransactionNumber = '{fSqlFormat(rs.Fields("TransactionNumber").Value)}' and Line = {fNum(rs.Fields("Line").Value)} and PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}' and Cashier = '{fSqlFormat(rs.Fields("Cashier").Value)}' and ItemCode = '{fSqlFormat(rs.Fields("ItemCode").Value)}' and Quantity = {fNum(rs.Fields("Quantity").Value)} ", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_PS_ItemsSold_Tmp 
                                                (TransactionNumber,
                                                Line,
                                                PSDate,
                                                [Counter],
                                                Cashier,
                                                ItemCode,
                                                Quantity,
                                                GrossSRP,
                                                Discount,
                                                Surcharge,
                                                TotalGross,
                                                TotalDiscount,
                                                TotalSurcharge,
                                                TotalNet,
                                                Location,
                                                Posted,
                                                POSTableKey)
                                                VALUES (      
                                                '{fSqlFormat(rs.Fields("TransactionNumber").Value)}',
                                                 {fNum(rs.Fields("Line").Value)},
                                                 {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                                '{fSqlFormat(rs.Fields("Counter").Value)}',
                                                '{fSqlFormat(rs.Fields("Cashier").Value)}',
                                                '{fSqlFormat(rs.Fields("ItemCode").Value)}',
                                                {fNum(rs.Fields("Quantity").Value)},
                                                {fNum(rs.Fields("GrossSRP").Value)},
                                                {fNum(rs.Fields("Discount").Value)},
                                                {fNum(rs.Fields("Surcharge").Value)},
                                                {fNum(rs.Fields("TotalGross").Value)},
                                                {fNum(rs.Fields("TotalDiscount").Value)},
                                                {fNum(rs.Fields("TotalSurcharge").Value)},
                                                {fNum(rs.Fields("TotalNet").Value)},
                                                '{fSqlFormat(rs.Fields("Location").Value)}',
                                                {fNum(rs.Fields("Posted").Value)},
                                                {fNum(rs.Fields("POSTableKey").Value)});"

                    ConnServer.Execute(strSQL)
                Else
                    Dim strSQL As String = $"UPDATE tbl_PS_ItemsSold_Tmp SET
                            GrossSRP          = {fNum(rs.Fields("GrossSRP").Value)},
                            Discount          = {fNum(rs.Fields("Discount").Value)},
                            Surcharge         = {fNum(rs.Fields("Surcharge").Value)},
                            TotalGross        = {fNum(rs.Fields("TotalGross").Value)},
                            TotalDiscount     = {fNum(rs.Fields("TotalDiscount").Value)},
                            TotalSurcharge    = {fNum(rs.Fields("TotalSurcharge").Value)},
                            TotalNet          = {fNum(rs.Fields("TotalNet").Value)},
                            Location          = '{fSqlFormat(rs.Fields("Location").Value)}',
                            Posted            = {fNum(rs.Fields("Posted").Value)},
                            POSTableKey       = {fNum(rs.Fields("POSTableKey").Value)}
                            WHERE  TransactionNumber = '{fSqlFormat(rs.Fields("TransactionNumber").Value)}' and 
                            Line = {fNum(rs.Fields("Line").Value)} and 
                            PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and 
                            [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}' and 
                            Cashier = '{fSqlFormat(rs.Fields("Cashier").Value)}' and 
                            ItemCode = '{fSqlFormat(rs.Fields("ItemCode").Value)}' and 
                            Quantity = {fNum(rs.Fields("Quantity").Value)};"

                    ConnServer.Execute(strSQL)

                End If

                rs.MoveNext()
            End While
        End If

    End Sub

    Public Sub Branch_Insert_tbl_PS_ItemsSold_Voided(pb As ProgressBar, l As Label)

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PS_ItemsSold_Voided ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_ItemsSold_Voided :" & pb.Maximum & "/" & pb.Value

                Application.DoEvents()
                Dim rx As New Recordset

                rx.Open($"select * from tbl_PS_ItemsSold_Voided WHERE TransactionNumber = '{fSqlFormat(rs.Fields("TransactionNumber").Value)}' and PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}' and Cashier = '{fSqlFormat(rs.Fields("Cashier").Value)}' and ItemCode = '{fSqlFormat(rs.Fields("ItemCode").Value)}' and Quantity = {fNum(rs.Fields("Quantity").Value)}", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_PS_ItemsSold_Voided 
                                                    (
                                                    TransactionNumber,
                                                    PSDate,
                                                    [Counter],
                                                    Cashier,
                                                    ItemCode,
                                                    Quantity,
                                                    GrossSRP,
                                                    [Discount],
                                                    Surcharge,
                                                    TotalGross,
                                                    TotalDiscount,
                                                    TotalSurcharge,
                                                    TotalNet,
                                                    Posted,
                                                    ViodedBy,
                                                    Location)
                                                VALUES (     
                                                '{fSqlFormat(rs.Fields("TransactionNumber").Value)}',
                                                 {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                                '{fSqlFormat(rs.Fields("Counter").Value)}',
                                                '{fSqlFormat(rs.Fields("Cashier").Value)}',
                                                '{fSqlFormat(rs.Fields("ItemCode").Value)}',
                                                {fNum(rs.Fields("Quantity").Value)},
                                                {fNum(rs.Fields("GrossSRP").Value)},
                                                {fNum(rs.Fields("Discount").Value)},
                                                {fNum(rs.Fields("Surcharge").Value)},
                                                {fNum(rs.Fields("TotalGross").Value)},
                                                {fNum(rs.Fields("TotalDiscount").Value)},
                                                {fNum(rs.Fields("TotalSurcharge").Value)},
                                                {fNum(rs.Fields("TotalNet").Value)},
                                                {fNum(rs.Fields("Posted").Value)},
                                                '{fSqlFormat(rs.Fields("ViodedBy").Value)}'
                                                '{fSqlFormat(rs.Fields("Location").Value)}'
                                                );"

                    ConnServer.Execute(strSQL)

                Else
                    Dim strSQL As String = $"UPDATE tbl_PS_ItemsSold_Voided SET
                            GrossSRP           = {fNum(rs.Fields("GrossSRP").Value)},
                            [Discount]         = {fNum(rs.Fields("Discount").Value)},
                            Surcharge          = {fNum(rs.Fields("Surcharge").Value)},
                            TotalGross         = {fNum(rs.Fields("TotalGross").Value)},
                            TotalDiscount      = {fNum(rs.Fields("TotalDiscount").Value)},
                            TotalSurcharge     = {fNum(rs.Fields("TotalSurcharge").Value)},
                            TotalNet           = {fNum(rs.Fields("TotalNet").Value)},
                            Posted             = {fNum(rs.Fields("Posted").Value)},
                            ViodedBy           = '{fSqlFormat(rs.Fields("ViodedBy").Value)}',
                            Location           = '{fSqlFormat(rs.Fields("Location").Value)}'
                            WHERE TransactionNumber = '{fSqlFormat(rs.Fields("TransactionNumber").Value)}' and PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}' and Cashier = '{fSqlFormat(rs.Fields("Cashier").Value)}' and ItemCode = '{fSqlFormat(rs.Fields("ItemCode").Value)}' and Quantity = {fNum(rs.Fields("Quantity").Value)};"

                    ConnServer.Execute(strSQL)

                End If

                rs.MoveNext()
            End While
        End If

    End Sub

    Public Sub Branch_Insert_tbl_PS_MiscPay_Tmp(pb As ProgressBar, l As Label)

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PS_MiscPay_Tmp  ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_MiscPay_Tmp  :" & pb.Maximum & "/" & pb.Value

                Application.DoEvents()

                Dim rx As New Recordset
                rx.Open($"SELECT TOP 1 * from tbl_PS_MiscPay_Tmp  WHERE TransactionNumber = '{fSqlFormat(rs.Fields("TransactionNumber").Value)}' and Line =  {fNum(rs.Fields("Line").Value)} and PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}' and Cashier = '{fSqlFormat(rs.Fields("Cashier").Value)}'", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_PS_MiscPay_Tmp 
                                                    (TransactionNumber,
                                                    Line,
                                                    PSDate,
                                                    [Counter],
                                                    Cashier,
                                                    Track1,
                                                    Track2,
                                                    Type,
                                                    Code,
                                                    BankKey,
                                                    TypePayment,
                                                    CardTerms,
                                                    [Account],
                                                    [Name],
                                                    Amount,
                                                    SSU,
                                                    Location,
                                                    Posted,
                                                    POSTableKey,
                                                    AmountAct,
                                                    [Tax],
                                                    BankComm)
                                                VALUES (      
                                                '{fSqlFormat(rs.Fields("TransactionNumber").Value)}',
                                                 {fNum(rs.Fields("Line").Value)},
                                                 {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                                '{fSqlFormat(rs.Fields("Counter").Value)}',
                                                '{fSqlFormat(rs.Fields("Cashier").Value)}',
                                                '{fSqlFormat(rs.Fields("Track1").Value)}',
                                                '{fSqlFormat(rs.Fields("Track2").Value)}',
                                                '{fSqlFormat(rs.Fields("Type").Value)}',
                                                '{fSqlFormat(rs.Fields("Code").Value)}',
                                                {fNum(rs.Fields("BankKey").Value)},
                                                {fNum(rs.Fields("TypePayment").Value)},
                                                '{fSqlFormat(rs.Fields("CardTerms").Value)}',
                                                '{fSqlFormat(rs.Fields("Account").Value)}',
                                                '{fSqlFormat(rs.Fields("Name").Value)}',
                                                 {fNum(rs.Fields("Amount").Value)},
                                                 {fNum(rs.Fields("SSU").Value)},
                                                '{fSqlFormat(rs.Fields("Location").Value)}',
                                                 {fNum(rs.Fields("Posted").Value)},
                                                 {fNum(rs.Fields("POSTableKey").Value)},
                                                 {fNum(rs.Fields("AmountAct").Value)},
                                                 {fNum(rs.Fields("Tax").Value)},
                                                 {fNum(rs.Fields("BankComm").Value)});"

                    ConnServer.Execute(strSQL)
                Else
                    Dim strSQL As String = $"UPDATE tbl_PS_MiscPay_Tmp SET
                            
                            Track1            = '{fSqlFormat(rs.Fields("Track1").Value)}',
                            Track2            = '{fSqlFormat(rs.Fields("Track2").Value)}',
                            [Type]            = '{fSqlFormat(rs.Fields("Type").Value)}',
                            Code              = '{fSqlFormat(rs.Fields("Code").Value)}',
                            BankKey           = {fNum(rs.Fields("BankKey").Value)},
                            TypePayment       = {fNum(rs.Fields("TypePayment").Value)},
                            CardTerms         = '{fSqlFormat(rs.Fields("CardTerms").Value)}',
                            [Account]         = '{fSqlFormat(rs.Fields("Account").Value)}',
                            [Name]            = '{fSqlFormat(rs.Fields("Name").Value)}',
                            Amount            = {fNum(rs.Fields("Amount").Value)},
                            SSU               = {fNum(rs.Fields("SSU").Value)},
                            Location          = '{fSqlFormat(rs.Fields("Location").Value)}',
                            Posted            = {fNum(rs.Fields("Posted").Value)},
                            POSTableKey       = {fNum(rs.Fields("POSTableKey").Value)},
                            AmountAct         = {fNum(rs.Fields("AmountAct").Value)},
                            [Tax]             = {fNum(rs.Fields("Tax").Value)},
                            BankComm          = {fNum(rs.Fields("BankComm").Value)}
                            WHERE TransactionNumber = '{fSqlFormat(rs.Fields("TransactionNumber").Value)}' and Line =  {fNum(rs.Fields("Line").Value)} and PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}' and Cashier = '{fSqlFormat(rs.Fields("Cashier").Value)}';"

                    ConnServer.Execute(strSQL)
                End If


                rs.MoveNext()
            End While
        End If

    End Sub
    Public Sub Branch_Insert_tbl_PS_MiscPay_Voided(pb As ProgressBar, l As Label)

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PS_MiscPay_Voided ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_MiscPay_Voided  :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim rx As New Recordset
                rx.Open($"SELECT * FROM tbl_PS_MiscPay_Voided  WHERE TransactionNumber = '{fSqlFormat(rs.Fields("TransactionNumber").Value)}' and PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}' and Cashier = '{fSqlFormat(rs.Fields("Cashier").Value)}' and [Code] = '{fSqlFormat(rs.Fields("Code").Value)}'", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_PS_MiscPay_Voided 
                                                    (TransactionNumber,
                                                    PSDate,
                                                    [Counter],
                                                    Cashier,
                                                    Track1,
                                                    Track2,
                                                    [Type],
                                                    [Code],
                                                    TypePayment,
                                                    BankKey,
                                                    CardTerms,
                                                    [Account],
                                                    [Name],
                                                    Amount,
                                                    SSU,
                                                    Posted,
                                                    ViodedBy,
                                                    Location)
                                                VALUES (      
                                                '{fSqlFormat(rs.Fields("TransactionNumber").Value)}',
                                                 {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                                '{fSqlFormat(rs.Fields("Counter").Value)}',
                                                '{fSqlFormat(rs.Fields("Cashier").Value)}',
                                                '{fSqlFormat(rs.Fields("Track1").Value)}',
                                                '{fSqlFormat(rs.Fields("Track2").Value)}',
                                                '{fSqlFormat(rs.Fields("Type").Value)}',
                                                '{fSqlFormat(rs.Fields("Code").Value)}',
                                                 {fNum(rs.Fields("TypePayment").Value)},
                                                 {fNum(rs.Fields("BankKey").Value)},         
                                                '{fSqlFormat(rs.Fields("CardTerms").Value)}',
                                                '{fSqlFormat(rs.Fields("Account").Value)}',
                                                '{fSqlFormat(rs.Fields("Name").Value)}',
                                                 {fNum(rs.Fields("Amount").Value)},
                                                 {fNum(rs.Fields("SSU").Value)},                                  
                                                 {fNum(rs.Fields("Posted").Value)},
                                                '{fSqlFormat(rs.Fields("ViodedBy").Value)}',
                                                '{fSqlFormat(rs.Fields("Location").Value)}'

                                                );"

                    ConnServer.Execute(strSQL)

                Else
                    Dim strSQL As String = $"UPDATE tbl_PS_MiscPay_Voided SET
                            PSDate      = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                            [Counter]   = '{fSqlFormat(rs.Fields("Counter").Value)}',
                            Cashier     = '{fSqlFormat(rs.Fields("Cashier").Value)}',
                            Track1      = '{fSqlFormat(rs.Fields("Track1").Value)}',
                            Track2      = '{fSqlFormat(rs.Fields("Track2").Value)}',
                            [Type]      = '{fSqlFormat(rs.Fields("Type").Value)}',
                            [Code]      = '{fSqlFormat(rs.Fields("Code").Value)}',
                            TypePayment = {fNum(rs.Fields("TypePayment").Value)},
                            BankKey     = {fNum(rs.Fields("BankKey").Value)},
                            CardTerms   = '{fSqlFormat(rs.Fields("CardTerms").Value)}',
                            [Account]   = '{fSqlFormat(rs.Fields("Account").Value)}',
                            [Name]      = '{fSqlFormat(rs.Fields("Name").Value)}',
                            Amount      = {fNum(rs.Fields("Amount").Value)},
                            SSU         = {fNum(rs.Fields("SSU").Value)},
                            Posted      = {fNum(rs.Fields("Posted").Value)},
                            ViodedBy    = '{fSqlFormat(rs.Fields("ViodedBy").Value)}',
                            Location    = '{fSqlFormat(rs.Fields("Location").Value)}'
                            WHERE   TransactionNumber = '{fSqlFormat(rs.Fields("TransactionNumber").Value)}' and 
                                    PSDate = {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())} and
                                    [Counter] = '{fSqlFormat(rs.Fields("Counter").Value)}' and 
                                    Cashier = '{fSqlFormat(rs.Fields("Cashier").Value)}' and 
                                    [Code] = '{fSqlFormat(rs.Fields("Code").Value)}';"

                    ConnServer.Execute(strSQL)

                End If

                rs.MoveNext()
            End While
        End If

    End Sub
    Public Sub Branch_Insert_tbl_PaidOutTransactions(pb As ProgressBar, l As Label)


        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PaidOutTransactions ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PaidOutTransactions  :" & pb.Maximum & "/" & pb.Value




                Application.DoEvents()

                Dim rx As New Recordset
                rx.Open($"SELECT TOP 1 * FROM tbl_PaidOutTransactions WHERE TransDate = {fDateIsEmpty(rs.Fields("TransDate").Value.ToString())} and 
                                                                            TransTime = '{fSqlFormat(rs.Fields("TransTime").Value)}' and  
                                                                            CashierCode = '{fSqlFormat(rs.Fields("CashierCode").Value)}' and 
                                                                            MachineNo = '{fSqlFormat(rs.Fields("MachineNo").Value)}' ", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then

                    Dim Series As Integer = 0
                    Dim D_Year As Integer = CType(rs.Fields("TransDate").Value, Date).Year

                    Dim s As String = $"SELECT TOP(1)  Series  FROM tbl_PaidOutTransactions  WHERE (YYear = {D_Year})  ORDER BY Series Desc "
                    Dim r As New Recordset
                    r.Open(s, ConnServer, CursorTypeEnum.adOpenStatic)
                    If r.RecordCount > 0 Then
                        Series = Val(r.Fields("Series").Value) + 1
                    Else
                        Series = 1
                    End If



                    Dim OldPK As Integer = fNum(rs.Fields("PaidOutPK").Value)

                    Dim strSQL As String = $"INSERT INTO tbl_PaidOutTransactions 
                                                    (   TransDate,
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
                                                VALUES ( {fDateIsEmpty(rs.Fields("TransDate").Value.ToString())},
                                                        '{fSqlFormat(rs.Fields("TransTime").Value)}',
                                                        '{ D_Year & Format(Series, "000000#")}',
                                                         {fNum(rs.Fields("OOrder").Value)},                                                  
                                                        '{fSqlFormat(rs.Fields("CashierCode").Value)}',
                                                        '{fSqlFormat(rs.Fields("CashierName").Value)}',
                                                        '{fSqlFormat(rs.Fields("CollectorCode").Value)}',
                                                        '{fSqlFormat(rs.Fields("CollectorName").Value)}',      
                                                        '{fSqlFormat(rs.Fields("MachineNo").Value)}',     
                                                         {fNum(rs.Fields("Total").Value)},
                                                         {fNum(rs.Fields("YYear").Value)},                                  
                                                         {Series},
                                                         {fNum(rs.Fields("IsPosted").Value)},
                                                         {fNum(rs.Fields("IsChecked").Value)},
                                                         {fNum(rs.Fields("Total_Previous").Value)},
                                                         {fNum(rs.Fields("SessionPK").Value)},
                                                         {fNum(rs.Fields("IsUsed").Value)}
                                                );"

                    ConnServer.Execute(strSQL)
                    Dim rsID As New Recordset
                    rsID = ConnServer.Execute("SELECT SCOPE_IDENTITY() AS NewID;")
                    Dim newPK As Integer = 0
                    If Not rsID.EOF Then
                        newPK = Convert.ToInt32(rsID.Fields("NewID").Value)
                        Branch_Insert_tbl_PaidOutTransactions_Det(OldPK, newPK) '  Must insert the details
                    End If
                    rsID.Close()

                End If


                rs.MoveNext()
            End While
        End If

    End Sub
    Private Sub Branch_Insert_tbl_PaidOutTransactions_Det(LocalPK As Integer, ServerPK As Integer)
        Dim rs1 As New ADODB.Recordset

        rs1.Open($"select * from tbl_PaidOutTransactions_Det  WHERE PaidOutPK = {LocalPK} ", ConnLocal, ADODB.CursorTypeEnum.adOpenStatic)
        If rs1.RecordCount > 0 Then
            While Not rs1.EOF
                Application.DoEvents()
                Dim rx As New Recordset
                rx.Open($"SELECT TOP 1 *  FROM tbl_PaidOutTransactions_Det  WHERE PaidOutPK = {ServerPK} and  DenomPK = {fNum(rs1.Fields("DenomPK").Value)} ", ConnServer, CursorTypeEnum.adOpenStatic)
                If rx.RecordCount = 0 Then
                    Dim strSQL As String = $"INSERT INTO tbl_PaidOutTransactions_Det 
                                                (  PaidOutPK,
                                                    DenomPK,
                                                    Qty,
                                                    DenomAmount,
                                                    Total,
                                                    STN_Qty,
                                                    STN_Amount,
                                                    IsChecked,
                                                    Old_Qty,
                                                    Old_Amount,
                                                    Remarks,
                                                    AdjustedBy,
                                                    WitnessedBy,
                                                    ApprovedBy,
                                                    Old_Qty_tmp,
                                                    DenomCode)
                                                VALUES ({ServerPK}, 
                                                        {fNum(rs1.Fields("DenomPK").Value)},   
                                                        {fNum(rs1.Fields("Qty").Value)},
                                                        {fNum(rs1.Fields("DenomAmount").Value)},
                                                        {fNum(rs1.Fields("Total").Value)},
                                                        {fNum(rs1.Fields("STN_Qty").Value)},
                                                        {fNum(rs1.Fields("STN_Amount").Value)},
                                                        {fNum(rs1.Fields("IsChecked").Value)},
                                                        {fNum(rs1.Fields("Old_Qty").Value)},
                                                        {fNum(rs1.Fields("Old_Amount").Value)},                                                       
                                                        '{fSqlFormat(rs1.Fields("Remarks").Value)}',     
                                                        '{fSqlFormat(rs1.Fields("AdjustedBy").Value)}',   
                                                        '{fSqlFormat(rs1.Fields("WitnessedBy").Value)}',                                               
                                                        '{fSqlFormat(rs1.Fields("ApprovedBy").Value)}',    
                                                         {fNum(rs1.Fields("Old_Qty_tmp").Value)},
                                                        '{fSqlFormat(rs1.Fields("DenomCode").Value)}');"

                    ConnServer.Execute(strSQL)
                End If

                rs1.MoveNext()
            End While
        End If

    End Sub

    Public Function GetBranchInfo() As Boolean
        Dim isHave As Boolean
        Try
            Dim rx As New Recordset
            rx.Open($"SELECT * FROM tbl_info WHERE [Counter] <> 'Main'", ConnLocal, CursorTypeEnum.adOpenStatic)
            If rx.RecordCount <> 0 Then
                isHave = True
            Else
                isHave = False
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error Upload", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        End Try


        GetBranchInfo = isHave

    End Function
End Module
