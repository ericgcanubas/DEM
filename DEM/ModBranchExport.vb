﻿Module ModBranchExport


    Public Sub Branch_CreateTable_tbl_GiftCert_List(pb As ProgressBar, l As Label, dt As DateTime)
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

            ConnLocal.Execute(createTableSql)
            Branch_Collect_tbl_GiftCert_List(pb, l, dt)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_GiftCert_List")
            Application.Exit()
        End Try
    End Sub

    Private Sub Branch_Collect_tbl_GiftCert_List(pb As ProgressBar, l As Label, dt As DateTime)

        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_GiftCert_List where DateAdded = {fDateIsEmpty(dt.ToShortDateString())}  and DateUsed is null ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
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
                ConnLocal.Execute(strSQL)
                rs.MoveNext()
            End While
        End If

    End Sub

    Public Sub Branch_CreateTable_tbl_VPlus_Codes(pb As ProgressBar, l As Label, dt As DateTime)
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

            ConnLocal.Execute(createTableSql)
            Branch_Collect_tbl_VPlus_Codes(pb, l, dt)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_VPlus_Codes")
            Application.Exit()
        End Try
    End Sub
    Private Sub Branch_Collect_tbl_VPlus_Codes(pb As ProgressBar, l As Label, dt As DateTime)

        Dim n As Integer = 0
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_VPlus_Codes where CreatedOn = {fDateIsEmpty(dt.ToShortDateString())} ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
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

                ConnLocal.Execute(strSQL)
                rs.MoveNext()
            End While
        End If

    End Sub

    Public Sub Branch_CreateTable_tbl_VPlus_Codes_Validity(pb As ProgressBar, l As Label, dt As DateTime)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_VPlus_Codes_Validity (
                                            Codes TEXT(16) NOT NULL,
                                            DateStarted DATETIME NOT NULL,
                                            DateExpired DATETIME NOT NULL,
                                            GracePeriod DATETIME NOT NULL
                                        );"

            ConnLocal.Execute(createTableSql)
            Branch_Collect_tbl_VPlus_Codes_Validity(pb, l, dt)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_VPlus_Codes_Validity")
            Application.Exit()
        End Try
    End Sub
    Private Sub Branch_Collect_tbl_VPlus_Codes_Validity(pb As ProgressBar, l As Label, dt As DateTime)

        rs = New ADODB.Recordset
        rs.Open($"select tbl_VPlus_Codes_Validity.* from tbl_VPlus_Codes_Validity join tbl_VPlus_Codes on tbl_VPlus_Codes.codes = tbl_VPlus_Codes_Validity.codes  where tbl_VPlus_Codes.CreatedOn = {fDateIsEmpty(dt.ToShortDateString())} ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
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

                ConnLocal.Execute(strSQL)
                rs.MoveNext()
            End While

        End If

    End Sub

    Public Sub Branch_CreateTable_tbl_PS_GT(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_PS_GT (
                                                [Counter] TEXT(3) NOT NULL,
                                                TransactionCount INTEGER NOT NULL,
                                                GrandTotal DOUBLE NOT NULL,
                                                ZZCount INTEGER NOT NULL,
                                                ResetCnt TEXT(20) NOT NULL,
                                                ResetTrans DOUBLE NOT NULL,
                                                InvoiceNumberOld TEXT(10) NOT NULL,
                                                InvoiceNumberCnt DOUBLE NOT NULL,
                                                InvoiceNumber TEXT(10) NOT NULL,
                                                RA DOUBLE NOT NULL,
                                                RACount INTEGER NOT NULL,
                                                Sales DOUBLE NOT NULL,
                                                SalesCount DOUBLE NOT NULL,
                                                Discount DOUBLE NOT NULL,
                                                Surcharge DOUBLE NOT NULL,
                                                TranCount INTEGER NOT NULL,
                                                Cash DOUBLE NOT NULL,
                                                CashCount INTEGER NOT NULL,
                                                Card DOUBLE NOT NULL,
                                                CardCount INTEGER NOT NULL,
                                                [GC] DOUBLE NOT NULL,
                                                GCCount INTEGER NOT NULL,
                                                IncentiveCard DOUBLE NOT NULL,
                                                IncentiveCardCount INTEGER NOT NULL,
                                                CreditMemo DOUBLE NOT NULL,
                                                CreditMemoCount INTEGER NOT NULL,
                                                CM_CashRefund DOUBLE NOT NULL,
                                                CM_CashRefundCount INTEGER NOT NULL,
                                                ATD DOUBLE NOT NULL,
                                                ATDCount INTEGER NOT NULL,
                                                VPlus DOUBLE NOT NULL,
                                                VPlusCount INTEGER NOT NULL,
                                                Misc DOUBLE NOT NULL,
                                                MiscCount INTEGER NOT NULL,
                                                SN TEXT(20) NOT NULL,
                                                PermitNo TEXT(50) NOT NULL,
                                                M_I_N TEXT(50) NOT NULL,
                                                Trans BYTE NOT NULL,
                                                Locked BYTE NOT NULL,
                                                VPlusCodeCount DOUBLE NOT NULL,
                                                Header1 TEXT(50) NOT NULL,
                                                Header2 TEXT(50) NOT NULL,
                                                Header3 TEXT(50) NOT NULL,
                                                TIN TEXT(50) NOT NULL,
                                                ForOfflineMode BYTE NOT NULL,
                                                CapableOffline BYTE NOT NULL,
                                                WithEJournal BYTE NOT NULL,
                                                BankCommission DOUBLE,
                                                SupplierName TEXT(50) NOT NULL,
                                                SupplierAddress1 TEXT(50) NOT NULL,
                                                SupplierAddress2 TEXT(50) NOT NULL,
                                                SupplierTIN TEXT(50) NOT NULL,
                                                SupplierAccreditationNo TEXT(50) NOT NULL,
                                                SupplierDateIssued TEXT(50) NOT NULL,
                                                SupplierValidUntil TEXT(50) NOT NULL,
                                                IsNewRegistered INTEGER NOT NULL,
                                                IsNew INTEGER NOT NULL,
                                                IsDisabled INTEGER NOT NULL,
                                                HomeCredit DOUBLE NOT NULL,
                                                HomeCreditCount DOUBLE NOT NULL,
                                                QRPay DOUBLE,
                                                QRPayCount DOUBLE
);

                                                "

            ConnLocal.Execute(createTableSql)
            Branch_Collect_tbl_PS_GT(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_PS_GT")
            Application.Exit()
        End Try
    End Sub
    Private Sub Branch_Collect_tbl_PS_GT(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 5

        rs = New ADODB.Recordset
        rs.Open($"SELECT * FROM tbl_PS_GT WHERE [Counter] = '{gbl_Counter}' ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_GT :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim strSQL As String = $"INSERT INTO tbl_PS_GT 
                                            (  [Counter],
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

                ConnLocal.Execute(strSQL)
                rs.MoveNext()
            End While

        End If

    End Sub
    Public Sub Branch_CreateTable_tbl_PS_GT_ZZ(pb As ProgressBar, l As Label)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_PS_GT_ZZ (
                                                [Counter] TEXT(3) NOT NULL,
                                                PSDate DATETIME NOT NULL,
                                                ZZCount BYTE NOT NULL
                                            );"

            ConnLocal.Execute(createTableSql)
            Branch_Collect_tbl_PS_GT_ZZ(pb, l)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_PS_GT_ZZ")
            Application.Exit()
        End Try
    End Sub
    Private Sub Branch_Collect_tbl_PS_GT_ZZ(pb As ProgressBar, l As Label)
        Dim year As Integer = Now.Year - 5

        rs = New ADODB.Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open($"select * from tbl_PS_GT_ZZ Where [Counter] = '{gbl_Counter}'", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_GT_ZZ :" & pb.Maximum & "/" & pb.Value
                Application.DoEvents()
                Dim strSQL As String = $"INSERT INTO tbl_PS_GT_ZZ 
                                            ([Counter],
                                            [PSDate],
                                            ZZCount)
                                    VALUES (
                                        '{fSqlFormat(rs.Fields("Counter").Value)}',
                                         {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                        {fNum(rs.Fields("ZZCount").Value)}
                                   );"

                ConnLocal.Execute(strSQL)
                rs.MoveNext()
            End While

        End If

    End Sub
    Public Sub Branch_CreateTable_tbl_PS_E_Journal(pb As ProgressBar, l As Label, dt As DateTime)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_PS_E_Journal (
                                            PK INTEGER PRIMARY KEY,
                                            PSNumber TEXT(15) NOT NULL,
                                            PSDate DATETIME NOT NULL,
                                            Cashier TEXT(3) NOT NULL,
                                            [Counter] TEXT(3) NOT NULL,
                                            Series TEXT(6) NOT NULL,
                                            ExactDate DATETIME NOT NULL,
                                            Amount DOUBLE NOT NULL,
                                            SRem TEXT(50),
                                            TotalQty DOUBLE NOT NULL,
                                            TotalSales DOUBLE NOT NULL,
                                            TotalDiscount DOUBLE NOT NULL,
                                            TotalGC DOUBLE NOT NULL,
                                            TotalCard DOUBLE NOT NULL,
                                            TotalVPlus DOUBLE NOT NULL,
                                            TotalATD DOUBLE NOT NULL,
                                            Location TEXT(1) NOT NULL,
                                            InvoiceNumber TEXT(15) NOT NULL,
                                            VatPercent TEXT(10) NOT NULL,
                                            VatSale DOUBLE NOT NULL,
                                            Vat DOUBLE NOT NULL,
                                            POSTableKey LONG NOT NULL,
                                            TotalIncentiveCard DOUBLE NOT NULL,
                                            IsZeroRated YESNO NOT NULL,
                                            TotalCreditMemo DOUBLE NOT NULL,
                                            TotalHomeCredit DOUBLE NOT NULL,
                                            TotalQRPay DOUBLE
                                        );"

            ConnLocal.Execute(createTableSql)
            Branch_Collect_tbl_PS_E_Journal(pb, l, dt)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_PS_E_Journal")
            Application.Exit()
        End Try
    End Sub
    Private Sub Branch_Collect_tbl_PS_E_Journal(pb As ProgressBar, l As Label, dt As DateTime)

        rs = New ADODB.Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open($"select * from tbl_PS_E_Journal where PsDate = {fDateIsEmpty(dt.ToShortDateString())}  and  [Counter] = '{gbl_Counter}' ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
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
                Dim strSQL As String = $"INSERT INTO tbl_PS_E_Journal  
                                            (PK,
                                            PSNumber,
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
                                    VALUES (
                                         {fNum(rs.Fields("PK").Value)},
                                        '{fSqlFormat(rs.Fields("PSNumber").Value)}',
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

                ConnLocal.Execute(strSQL)
                rs.MoveNext()
            End While
        End If
    End Sub

    Public Sub Branch_CreateTable_tbl_PS_E_Journal_Detail(pb As ProgressBar, l As Label, dt As DateTime)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_PS_E_Journal_Detail (
                                            PK INTEGER PRIMARY KEY,
                                            TransactionNumber TEXT(15) NOT NULL,
                                            PSDate DATETIME NOT NULL,
                                            [Counter] TEXT(3) NOT NULL,
                                            Cashier TEXT(3) NOT NULL,
                                            ItemCode TEXT(12) NOT NULL,
                                            ItemDescription TEXT(50) NOT NULL,
                                            Quantity DOUBLE NOT NULL,
                                            GrossSRP DOUBLE NOT NULL,
                                            Discount DOUBLE NOT NULL,
                                            Surcharge DOUBLE NOT NULL,
                                            TotalGross DOUBLE NOT NULL,
                                            TotalDiscount DOUBLE NOT NULL,
                                            TotalSurcharge DOUBLE NOT NULL,
                                            TotalNet DOUBLE NOT NULL,
                                            Location TEXT(1) NOT NULL,
                                            POSTableKey LONG NOT NULL
                                        );"

            ConnLocal.Execute(createTableSql)
            Branch_Collect_tbl_PS_E_Journal_Detail(pb, l, dt)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_PS_E_Journal_Detail")
            Application.Exit()
        End Try
    End Sub
    Private Sub Branch_Collect_tbl_PS_E_Journal_Detail(pb As ProgressBar, l As Label, dt As DateTime)
        Dim year As Integer = Now.Year - 1
        Dim toDate As String = Now.Date.ToShortDateString()
        Dim FromDate As String = Now.Date.AddYears(-1).ToShortDateString()
        rs = New ADODB.Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open($"select * from tbl_PS_E_Journal_Detail  where PsDate = {fDateIsEmpty(dt.ToShortDateString())}  and [Counter] = '{gbl_Counter}' ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
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
                Dim strSQL As String = $"INSERT INTO tbl_PS_E_Journal_Detail 
                                            (PK,
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
                                            VALUES ({fNum(rs.Fields("PK").Value)},
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

                ConnLocal.Execute(strSQL)
                rs.MoveNext()
            End While

        End If

    End Sub


    Public Sub Branch_CreateTable_tbl_PS_EmployeeATD(pb As ProgressBar, l As Label, dt As DateTime)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_PS_EmployeeATD (
                                                PK INTEGER PRIMARY KEY,
                                                TransactionNumber TEXT(15) NOT NULL,
                                                PSDate DATETIME NOT NULL,
                                                [Counter] TEXT(3) NOT NULL,
                                                Cashier TEXT(3) NOT NULL,
                                                ATDNumber TEXT(50) NOT NULL,
                                                EmpNo DOUBLE NOT NULL,
                                                Amount DOUBLE NOT NULL
                                            );"

            ConnLocal.Execute(createTableSql)
            Branch_Collect_tbl_PS_EmployeeATD(pb, l, dt)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_PS_EmployeeATD")
            Application.Exit()
        End Try
    End Sub
    Private Sub Branch_Collect_tbl_PS_EmployeeATD(pb As ProgressBar, l As Label, dt As DateTime)

        Dim n As Integer = 0
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PS_EmployeeATD where PSDate = {fDateIsEmpty(dt.ToShortDateString())} ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_EmployeeATD :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
                Dim strSQL As String = $"INSERT INTO tbl_PS_EmployeeATD 
                                                (PK,
                                                TransactionNumber,
                                                PSDate,
                                                [Counter],
                                                Cashier,
                                                ATDNumber,
                                                EmpNo,
                                                Amount)
                                                VALUES ({fNum(rs.Fields("PK").Value)},      
                                                '{fSqlFormat(rs.Fields("TransactionNumber").Value)}',
                                                 {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                                '{fSqlFormat(rs.Fields("Counter").Value)}',
                                                '{fSqlFormat(rs.Fields("Cashier").Value)}',
                                                '{fSqlFormat(rs.Fields("ATDNumber").Value)}',
                                                {fNum(rs.Fields("EmpNo").Value)},
                                                {fNum(rs.Fields("Amount").Value)});"

                ConnLocal.Execute(strSQL)
                rs.MoveNext()
            End While
        End If

    End Sub
    Public Sub Branch_CreateTable_tbl_GiftCert_Payment(pb As ProgressBar, l As Label, dt As DateTime)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_GiftCert_Payment (
                                            PK INTEGER PRIMARY KEY,
                                            PSNumber TEXT(20) NOT NULL,
                                            PSDate DATETIME NOT NULL,
                                            [Counter] TEXT(3) NOT NULL,
                                            Cashier TEXT(3) NOT NULL,
                                            GCNumber TEXT(50) NOT NULL,
                                            GCAmount DOUBLE NOT NULL,
                                            Posted BYTE NOT NULL
                                        );"

            ConnLocal.Execute(createTableSql)
            Branch_Collect_tbl_GiftCert_Payment(pb, l, dt)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_GiftCert_Payment")
            Application.Exit()
        End Try
    End Sub
    Private Sub Branch_Collect_tbl_GiftCert_Payment(pb As ProgressBar, l As Label, dt As DateTime)

        Dim n As Integer = 0
        rs = New ADODB.Recordset
        rs.Open($"SELECT * from tbl_GiftCert_Payment where PSDate = {fDateIsEmpty(dt.ToShortDateString())} ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_GiftCert_Payment :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
                Dim strSQL As String = $"INSERT INTO tbl_GiftCert_Payment 
                                                (PK,
                                                PSNumber,
                                                PSDate,
                                                [Counter],
                                                Cashier,
                                                GCNumber,
                                                GCAmount,
                                                Posted)
                                                VALUES ({fNum(rs.Fields("PK").Value)},      
                                                '{fSqlFormat(rs.Fields("PSNumber").Value)}',
                                                 {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                                '{fSqlFormat(rs.Fields("Counter").Value)}',
                                                '{fSqlFormat(rs.Fields("Cashier").Value)}',
                                                '{fSqlFormat(rs.Fields("GCNumber").Value)}',
                                                {fNum(rs.Fields("GCAmount").Value)},
                                                {fNum(rs.Fields("Posted").Value)});"

                ConnLocal.Execute(strSQL)
                rs.MoveNext()
            End While
        End If

    End Sub
    Public Sub Branch_CreateTable_tbl_VPlus_Purchases_Points(pb As ProgressBar, l As Label, dt As DateTime)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_VPlus_Purchases_Points (
                                            PK INTEGER PRIMARY KEY,
                                            TransactionNo TEXT(15) NOT NULL,
                                            VDate DATETIME NOT NULL,
                                            VPlusCodes TEXT(16) NOT NULL,
                                            [Counter] TEXT(3) NOT NULL,
                                            Cashier TEXT(3) NOT NULL,
                                            Cash DOUBLE NOT NULL,
                                            Card DOUBLE NOT NULL,
                                            [GC] DOUBLE NOT NULL,
                                            [ATD] DOUBLE NOT NULL,
                                            PointsPay DOUBLE NOT NULL,
                                            [Location] TEXT(1) NOT NULL,
                                            Posted BYTE NOT NULL
                                        );"

            ConnLocal.Execute(createTableSql)
            Branch_Collect_tbl_VPlus_Purchases_Points(pb, l, dt)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_VPlus_Purchases_Points")
            Application.Exit()
        End Try
    End Sub
    Private Sub Branch_Collect_tbl_VPlus_Purchases_Points(pb As ProgressBar, l As Label, dt As DateTime)

        Dim n As Integer = 0
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_VPlus_Purchases_Points where VDate = {fDateIsEmpty(dt.ToShortDateString())} ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_VPlus_Purchases_Points :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
                Dim strSQL As String = $"INSERT INTO tbl_VPlus_Purchases_Points 
                                                (PK,
                                                TransactionNo,
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
                                                VALUES ({fNum(rs.Fields("PK").Value)},      
                                                '{fSqlFormat(rs.Fields("TransactionNo").Value)}',
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
                rs.MoveNext()
            End While
        End If

    End Sub


    Public Sub Branch_CreateTable_tbl_PS_Tmp(pb As ProgressBar, l As Label, dt As DateTime)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_PS_Tmp (
                                            PK INTEGER PRIMARY KEY,
                                            PSNumber TEXT(15) NOT NULL,
                                            PSDate DATETIME NOT NULL,
                                            PSDateActual DATETIME NOT NULL,
                                            Cashier TEXT(3) NOT NULL,
                                            [Counter] TEXT(3) NOT NULL,
                                            [Series] TEXT(6) NOT NULL,
                                            ExactDate DATETIME NOT NULL,
                                            Amount DOUBLE NOT NULL,
                                            SRem TEXT(50) NOT NULL,
                                            TotalQty DOUBLE NOT NULL,
                                            TotalSales DOUBLE NOT NULL,
                                            TotalDiscount DOUBLE NOT NULL,
                                            TotalGC DOUBLE NOT NULL,
                                            TotalCard DOUBLE NOT NULL,
                                            TotalVPlus DOUBLE NOT NULL,
                                            TotalATD DOUBLE NOT NULL,
                                            Location TEXT(1) NOT NULL,
                                            InvoiceNumber TEXT(50),
                                            Posted BYTE NOT NULL,
                                            POSTableKey INTEGER NOT NULL,
                                            TotalIncentiveCard DOUBLE NOT NULL,
                                            IsZeroRated YESNO NOT NULL,
                                            TotalCreditMemo DOUBLE NOT NULL,
                                            TotalHomeCredit DOUBLE NOT NULL,
                                            TotalQRPay DOUBLE
                                        );"

            ConnLocal.Execute(createTableSql)
            Branch_Collect_tbl_PS_Tmp(pb, l, dt)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_PS_Tmp")
            Application.Exit()
        End Try
    End Sub
    Private Sub Branch_Collect_tbl_PS_Tmp(pb As ProgressBar, l As Label, dt As DateTime)

        Dim n As Integer = 0
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PS_Tmp where PSDate = {fDateIsEmpty(dt.ToShortDateString())} ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_Tmp :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
                Dim strSQL As String = $"INSERT INTO tbl_PS_Tmp 
                                                (PK,
                                                PSNumber,
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
                                                VALUES ({fNum(rs.Fields("PK").Value)},      
                                                '{fSqlFormat(rs.Fields("PSNumber").Value)}',
                                                 {fDateIsEmpty(rs.Fields("PSDate").Value.ToString())},
                                                 {fDateIsEmpty(rs.Fields("PSDateActual").Value.ToString())},
                                                '{fSqlFormat(rs.Fields("Cashier").Value)}',
                                                '{fSqlFormat(rs.Fields("Counter").Value)}',         
                                                '{fSqlFormat(rs.Fields("Series").Value)}',
                                                {fDateIsEmpty(rs.Fields("PSDateActual").Value.ToString())},
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

                ConnLocal.Execute(strSQL)
                rs.MoveNext()
            End While
        End If

    End Sub


    Public Sub Branch_CreateTable_tbl_PS_ItemsSold_Tmp(pb As ProgressBar, l As Label, dt As DateTime)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_PS_ItemsSold_Tmp (
                                            PK INTEGER PRIMARY KEY,
                                            TransactionNumber TEXT(15) NOT NULL,
                                            Line INTEGER NOT NULL,
                                            PSDate DATETIME NOT NULL,
                                            [Counter] TEXT(3) NOT NULL,
                                            Cashier TEXT(3) NOT NULL,
                                            ItemCode TEXT(12) NOT NULL,
                                            Quantity DOUBLE NOT NULL,
                                            GrossSRP DOUBLE NOT NULL,
                                            Discount DOUBLE NOT NULL,
                                            Surcharge DOUBLE NOT NULL,
                                            TotalGross DOUBLE NOT NULL,
                                            TotalDiscount DOUBLE NOT NULL,
                                            TotalSurcharge DOUBLE NOT NULL,
                                            TotalNet DOUBLE NOT NULL,
                                            Location TEXT(1) NOT NULL,
                                            Posted BYTE NOT NULL,
                                            POSTableKey INTEGER NOT NULL
                                        );"

            ConnLocal.Execute(createTableSql)
            Branch_Collect_tbl_PS_ItemsSold_Tmp(pb, l, dt)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_PS_ItemsSold_Tmp")
            Application.Exit()
        End Try
    End Sub
    Private Sub Branch_Collect_tbl_PS_ItemsSold_Tmp(pb As ProgressBar, l As Label, dt As DateTime)

        Dim n As Integer = 0
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PS_ItemsSold_Tmp where PSDate = {fDateIsEmpty(dt.ToShortDateString())} ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_ItemsSold_Tmp :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
                Dim strSQL As String = $"INSERT INTO tbl_PS_ItemsSold_Tmp 
                                                (PK,
                                                TransactionNumber,
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
                                                VALUES ({fNum(rs.Fields("PK").Value)},      
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
                                                {fNum(rs.Fields("POSTableKey").Value)}
                                                );"

                ConnLocal.Execute(strSQL)
                rs.MoveNext()
            End While
        End If

    End Sub

    Public Sub Branch_CreateTable_tbl_PS_ItemsSold_Voided(pb As ProgressBar, l As Label, dt As DateTime)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_PS_ItemsSold_Voided (
                                            PK INTEGER PRIMARY KEY,
                                            TransactionNumber TEXT(15) NOT NULL,
                                            PSDate DATETIME NOT NULL,
                                            [Counter] TEXT(3) NOT NULL,
                                            Cashier TEXT(3) NOT NULL,
                                            ItemCode TEXT(12) NOT NULL,
                                            Quantity DOUBLE NOT NULL,
                                            GrossSRP DOUBLE NOT NULL,
                                            [Discount] DOUBLE NOT NULL,
                                            Surcharge DOUBLE NOT NULL,
                                            TotalGross DOUBLE NOT NULL,
                                            TotalDiscount DOUBLE NOT NULL,
                                            TotalSurcharge DOUBLE NOT NULL,
                                            TotalNet DOUBLE NOT NULL,
                                            Posted BYTE NOT NULL,
                                            ViodedBy TEXT(50) NOT NULL,
                                            Location TEXT(1) NOT NULL
                                        );
                                        "

            ConnLocal.Execute(createTableSql)
            Branch_Collect_tbl_PS_ItemsSold_Voided(pb, l, dt)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_PS_ItemsSold_Voided")
            Application.Exit()
        End Try
    End Sub
    Private Sub Branch_Collect_tbl_PS_ItemsSold_Voided(pb As ProgressBar, l As Label, dt As DateTime)

        Dim n As Integer = 0
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PS_ItemsSold_Voided where PSDate = {fDateIsEmpty(dt.ToShortDateString())} ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_ItemsSold_Voided :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
                Dim strSQL As String = $"INSERT INTO tbl_PS_ItemsSold_Voided 
                                                    (PK,
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
                                                VALUES ({fNum(rs.Fields("PK").Value)},      
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

                ConnLocal.Execute(strSQL)
                rs.MoveNext()
            End While
        End If

    End Sub

    Public Sub Branch_CreateTable_tbl_PS_MiscPay_Tmp(pb As ProgressBar, l As Label, dt As DateTime)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_PS_MiscPay_Tmp (
                                            PK INTEGER PRIMARY KEY,
                                            TransactionNumber TEXT(15) NOT NULL,
                                            Line INTEGER NOT NULL,
                                            PSDate DATETIME NOT NULL,
                                            Counter TEXT(3) NOT NULL,
                                            Cashier TEXT(3) NOT NULL,
                                            Track1 TEXT(50),
                                            Track2 TEXT(50),
                                            Type TEXT(50) NOT NULL,
                                            Code TEXT(50) NOT NULL,
                                            BankKey INTEGER,
                                            TypePayment INTEGER NOT NULL,
                                            CardTerms TEXT(50) NOT NULL,
                                            [Account] TEXT(50),
                                            [Name] TEXT(50),
                                            Amount DOUBLE NOT NULL,
                                            SSU DOUBLE NOT NULL,
                                            Location TEXT(1) NOT NULL,
                                            Posted BYTE NOT NULL,
                                            POSTableKey INTEGER NOT NULL,
                                            AmountAct DOUBLE NOT NULL,
                                            [Tax] DOUBLE NOT NULL,
                                            BankComm DOUBLE NOT NULL
                                        );

                                        "

            ConnLocal.Execute(createTableSql)
            Branch_Collect_tbl_PS_MiscPay_Tmp(pb, l, dt)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_PS_MiscPay_Tmp ")
            Application.Exit()
        End Try
    End Sub
    Private Sub Branch_Collect_tbl_PS_MiscPay_Tmp(pb As ProgressBar, l As Label, dt As DateTime)

        Dim n As Integer = 0
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PS_MiscPay_Tmp  where PSDate = {fDateIsEmpty(dt.ToShortDateString())} ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_MiscPay_Tmp  :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
                Dim strSQL As String = $"INSERT INTO tbl_PS_MiscPay_Tmp 
                                                    (PK,
                                                    TransactionNumber,
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
                                                    BankComm
)
                                                VALUES ({fNum(rs.Fields("PK").Value)},      
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
                                                 {fNum(rs.Fields("BankComm").Value)}


                                                );"

                ConnLocal.Execute(strSQL)
                rs.MoveNext()
            End While
        End If

    End Sub


    Public Sub Branch_CreateTable_tbl_PS_MiscPay_Voided(pb As ProgressBar, l As Label, dt As Date)
        Try
            Dim createTableSql As String = "CREATE TABLE tbl_PS_MiscPay_Voided (
                                            PK INTEGER PRIMARY KEY,
                                            TransactionNumber TEXT(15) NOT NULL,
                                            PSDate DATETIME NOT NULL,
                                            [Counter] TEXT(3) NOT NULL,
                                            Cashier TEXT(3) NOT NULL,
                                            Track1 TEXT(50),
                                            Track2 TEXT(50),
                                            [Type] TEXT(50),
                                            [Code] TEXT(50),
                                            TypePayment INTEGER NOT NULL,
                                            BankKey INTEGER,
                                            CardTerms TEXT(50) NOT NULL,
                                            [Account] TEXT(50),
                                            [Name] TEXT(50),
                                            Amount DOUBLE NOT NULL,
                                            SSU DOUBLE NOT NULL,
                                            Posted BYTE NOT NULL,
                                            ViodedBy TEXT(50) NOT NULL,
                                            Location TEXT(1) NOT NULL
                                        );


                                        "

            ConnLocal.Execute(createTableSql)
            Branch_Collect_tbl_PS_MiscPay_Voided(pb, l, dt)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "tbl_PS_MiscPay_Voided  ")
            Application.Exit()
        End Try
    End Sub
    Private Sub Branch_Collect_tbl_PS_MiscPay_Voided(pb As ProgressBar, l As Label, dt As Date)

        Dim n As Integer = 0
        rs = New ADODB.Recordset
        rs.Open($"select * from tbl_PS_MiscPay_Voided  where PSDate = {fDateIsEmpty(dt.ToShortDateString())} ", ConnServer, ADODB.CursorTypeEnum.adOpenStatic)
        pb.Maximum = rs.RecordCount
        pb.Value = 0
        pb.Minimum = 0
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                pb.Value = pb.Value + 1
                l.Text = "tbl_PS_MiscPay_Voided  :" & pb.Maximum & "/" & pb.Value
                If n > 10000 Then
                    n = 0
                    Application.DoEvents()
                End If
                Dim strSQL As String = $"INSERT INTO tbl_PS_MiscPay_Voided 
                                                    (PK,
                                                    TransactionNumber,
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
                                                    Location
)
                                                VALUES ({fNum(rs.Fields("PK").Value)},      
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

                ConnLocal.Execute(strSQL)
                rs.MoveNext()
            End While
        End If

    End Sub


End Module
