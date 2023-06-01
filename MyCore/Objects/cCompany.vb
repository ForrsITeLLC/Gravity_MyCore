Imports MyCore.Data

Public Class cCompany

    Dim _CustomerNo As String = ""
    Dim _IsVendor As Boolean = False
    Dim _TermsName As String = ""
    Dim _DaysToDue As Integer = 30
    Dim _DaysToLate As Integer = 31

    Public Name As String = ""
    Public Address1 As String = ""
    Public Address2 As String = ""
    Public City As String = ""
    Public State As String = ""
    Public Zip As String = ""
    Public Country As String = ""
    Public Alert As Boolean = False
    Public AlertText As String = ""
    Public Terms As Integer = 0
    Public Type As Integer = 1
    Public ParentNo As String = ""
    Public Phone As String = ""
    Public Fax As String = ""
    Public WebSite As String = ""
    Public SalesTerritory As Integer = 0
    Public OfficeNo As Integer = 0
    Public ServiceArea As Integer = 0
    Public County As String = ""
    Public OurCustomerNo As String = ""
    Public TickerSymbol As String = ""
    Public Industry As Integer = 0
    Public GenerateCerts As Boolean = True
    Public ISO As Boolean = True
    Public CertTemplate As Integer = 0
    Public CertTemplateName As String = ""
    Public MarketingExclude As Boolean = False
    Public Notes As String = ""
    Public Password As String = ""
    Public Inactive As Boolean = False
    Public ServiceNote As String = ""
    Public RoundtripMiles As Integer = 0

    Public BillingName As String = ""
    Public BillingAddress1 As String = ""
    Public BillingAddress2 As String = ""
    Public BillingCity As String = ""
    Public BillingState As String = ""
    Public BillingZip As String = ""
    Public BillingCountry As String = ""
    Public BillingPhone As String = ""
    Public BillingFax As String = ""
    Public AccountNo As String = ""
    Public RateTable As Integer = 0
    Public TaxNo As String = ""
    Public TaxExemptThrough As Date = Nothing
    Public Ownership As String = ""
    Public TaxCode As Integer = 0
    Public AccountingCustomerNo As String = ""
    Public APEmailAddress As String = ""
    Public PreferredDeliveryMethod As Integer = 1  ' Mail = 1   Fax = 2   Email = 3
    Public BlanketPO As String = ""

    Public CustomText1 As String = ""
    Public CustomText2 As String = ""
    Public CustomText3 As String = ""

    Public LastUpdatedBy As String = ""
    Public CreatedBy As String = ""

    Public AcctRef As String = Nothing
    Public AcctUpload As Boolean = Nothing
    Public TermsRef As String = Nothing
    Public SalesTaxRef As String = Nothing

    Public PrintInvoiceSetting As Integer = 0
    Public InvoiceLineItemSetting As Integer = 0

    Dim _DaysSinceLastInvoice As Integer = -1
    Dim _DaysSinceFirstInvoice As Integer = -1
    Dim _DaysSinceOldestUnpaidInvoice As Integer = -1

    Dim _Office As DataTable
    Dim _TaxCode As DataTable
    Dim _RateTable As DataTable
    Dim _SalesTerritories As DataTable
    Dim _ServiceAreas As DataTable
    Dim _Terms As DataTable
    Dim _Countries As DataTable
    Dim _States As DataTable
    Dim _Industries As DataTable

    Dim Database As MyCore.Data.EasySql
    Dim Settings As cSettings

    Public Event Reload()
    Public Event AfterSave(ByVal Company As cCompany)

    Public ReadOnly Property TermsName() As String
        Get
            Return Me._TermsName
        End Get
    End Property

    Public ReadOnly Property DaysToLate() As Integer
        Get
            Return Me._DaysToLate
        End Get
    End Property

    Public ReadOnly Property DaysToDue() As Integer
        Get
            Return Me._DaysToDue
        End Get
    End Property

    Private ReadOnly Property BaseZip() As String
        Get
            If Me.Zip.Contains("-") Then
                Return Me.Zip.Substring(0, Me.Zip.IndexOf("-"))
            Else
                Return Me.Zip
            End If
        End Get
    End Property

    Public ReadOnly Property Zone(Optional ByVal Office As String = "") As Integer
        Get
            If Office.Length = 0 Then
                Office = Me.OfficeNo
            End If
            Dim Sql As String = "SELECT zone FROM zip_to_zone"
            Sql &= " WHERE office=" & Me.Database.Escape(Office)
            Sql &= " AND zip=" & Me.BaseZip
            Dim Result As DataTable = Me.Database.GetAll(Sql)
            If Me.Database.LastQuery.RowsReturned > 0 Then
                Return Result.Rows(0).Item(0)
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property IsContractCustomer() As Boolean
        Get
            Dim Months As Integer = Me.Settings.GetValue("Minimum Frequency for Contract Rate", 6)
            If Months >= 0 Then
                Dim Sql As String = "SELECT COUNT(id) AS num FROM cal_agreement"
                Sql &= " WHERE ship_to_no=" & Me.Database.Escape(Me.CustomerNo)
                Sql &= " AND canceled=0"
                If Months > 0 Then
                    Sql &= " AND frequency_months <= " & Months
                End If
                If Me.Database.GetOne(Sql) > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property IsTaxable() As Boolean
        Get
            If Me.TaxExemptThrough <> Nothing Then
                If Me.TaxExemptThrough >= Today Then
                    Return False
                End If
            End If
            Return True
        End Get
    End Property

    Public ReadOnly Property DaysSinceLastInvoice() As Integer
        Get
            Return Me._DaysSinceLastInvoice
        End Get
    End Property

    Public ReadOnly Property DaysSinceFirstInvoice() As Integer
        Get
            Return Me._DaysSinceFirstInvoice
        End Get
    End Property

    Public ReadOnly Property DaysSinceOldestUnpaidInvoice() As Integer
        Get
            Return Me._DaysSinceOldestUnpaidInvoice
        End Get
    End Property

    Public ReadOnly Property Offices() As DataTable
        Get
            Return Me._Office
        End Get
    End Property

    Public ReadOnly Property SalesTerritories() As DataTable
        Get
            Return Me._SalesTerritories
        End Get
    End Property

    Public ReadOnly Property ServiceAreas() As DataTable
        Get
            Return Me._ServiceAreas
        End Get
    End Property

    Public ReadOnly Property TermsList() As DataTable
        Get
            Return Me._Terms
        End Get
    End Property

    Public ReadOnly Property Countries() As DataTable
        Get
            Return Me._Countries
        End Get
    End Property

    Public ReadOnly Property States() As DataTable
        Get
            Return Me._States
        End Get
    End Property

    Public ReadOnly Property RateTables() As DataTable
        Get
            Return Me._RateTable
        End Get
    End Property

    Public ReadOnly Property TaxCodes() As DataTable
        Get
            Return Me._TaxCode
        End Get
    End Property

    Public ReadOnly Property CustomerNo() As String
        Get
            Return Me._CustomerNo
        End Get
    End Property

    Public ReadOnly Property Industries() As DataTable
        Get
            Return Me._Industries
        End Get
    End Property

    Public Property IsVendor() As Boolean
        Get
            Return Me._IsVendor
        End Get
        Set(ByVal value As Boolean)
            Me._IsVendor = value
        End Set
    End Property


    Public ReadOnly Property DefaultTechnician() As String
        Get
            Return Me.Database.GetOne("SELECT technician FROM service_area WHERE id=" & Me.ServiceArea)
        End Get
    End Property

    Public ReadOnly Property OutstandingInvoiceTotal() As Double
        Get
            ' Total of invoice line items not marked paid
            Dim Sql As String = ""
            Sql &= " SELECT "
            Sql &= " SUM([balance]) AS [total_balance]"
            Sql &= " FROM ("
            Sql &= "    SELECT  [ivi_ino], "
            Sql &= "    ISNULL(SUM([ivi_qua] * [ivi_cost]), 0)  + ISNULL(AVG([inv_shch]), 0)  "
            Sql &= "    + MAX(sales_tax) - ISNULL((SELECT SUM(pgl_amt)"
            Sql &= "    FROM PYMNTGLD WHERE pgl_ino=i.ivi_ino GROUP BY pgl_ino), 0)"
            Sql &= "    AS [balance]"
            Sql &= "    FROM INVITEMS i"
            Sql &= "    INNER JOIN INVOICE inv ON [ivi_ino]=[inv_ino]"
            Sql &= "    WHERE [ivi_ino] IN (SELECT [inv_ino] FROM INVOICE"
            Sql &= "    WHERE refund=0 AND inv_paid=0 AND inv_no=@cst_no AND voided=0) "
            Sql &= "    GROUP BY [ivi_ino]"
            Sql &= " ) AS i"
            'Sql &= " GROUP BY ivi_ino"
            Sql = Sql.Replace("@cst_no", Me.Database.Escape(Me.CustomerNo))
            Dim Value As Double = 0
            Try
                Value = Me.Database.GetOne(Sql)
            Catch
                ' Nothing
            End Try
            Return Value
        End Get
    End Property

    Public ReadOnly Property OutstandingInvoiceReport() As DataTable
        Get
            Dim Sql As String = "    SELECT  [ivi_ino] AS invoice_no, "
            Sql &= "    ISNULL(SUM([ivi_qua] * [ivi_cost]), 0)  + ISNULL(AVG([inv_shch]), 0)  "
            Sql &= "    + MAX(sales_tax) - ISNULL((SELECT SUM(pgl_amt)"
            Sql &= "    FROM PYMNTGLD WHERE pgl_ino=i.ivi_ino GROUP BY pgl_ino), 0)"
            Sql &= "    AS [balance]"
            Sql &= "    FROM INVITEMS i"
            Sql &= "    INNER JOIN INVOICE inv ON [ivi_ino]=[inv_ino]"
            Sql &= "    WHERE [ivi_ino] IN (SELECT [inv_ino] FROM INVOICE"
            Sql &= "    WHERE refund=0 AND inv_paid=0 AND inv_no=@cst_no AND voided=0) "
            Sql &= "    GROUP BY [ivi_ino]"
            Return Me.Database.GetAll(Sql)
        End Get
    End Property

    Public ReadOnly Property OutstandingRefundsTotal() As Double
        Get
            ' Total of invoice line items not marked paid
            Dim Sql As String = ""
            Sql &= " SELECT "
            Sql &= " SUM([balance]) AS [total_balance]"
            Sql &= " FROM ("
            Sql &= "    SELECT  [ivi_ino], "
            Sql &= "    ISNULL(SUM([ivi_qua] * [ivi_cost]), 0)  + ISNULL(AVG([inv_shch]), 0)  "
            Sql &= "    + MAX(sales_tax) - ISNULL((SELECT SUM(pgl_amt)"
            Sql &= "    FROM PYMNTGLD WHERE pgl_ino=i.ivi_ino GROUP BY pgl_ino), 0)"
            Sql &= "    AS [balance]"
            Sql &= "    FROM INVITEMS i"
            Sql &= "    INNER JOIN INVOICE inv ON [ivi_ino]=[inv_ino]"
            Sql &= "    WHERE [ivi_ino] IN (SELECT [inv_ino] FROM INVOICE"
            Sql &= "    WHERE refund=1 AND inv_paid=0 AND inv_no=@cst_no AND voided=0) "
            Sql &= "    GROUP BY [ivi_ino]"
            Sql &= " ) AS i"
            'Sql &= " GROUP BY ivi_ino"
            Sql = Sql.Replace("@cst_no", Me.Database.Escape(Me.CustomerNo))
            Dim Value As Double = 0
            Try
                Value = Me.Database.GetOne(Sql)
            Catch
                ' Nothing
            End Try
            Return Value
        End Get
    End Property

    Public ReadOnly Property OutstandingRefundsReport() As DataTable
        Get
            Dim Sql As String = "    SELECT  [ivi_ino] AS invoice_no, "
            Sql &= "    ISNULL(SUM([ivi_qua] * [ivi_cost]), 0)  + ISNULL(AVG([inv_shch]), 0)  "
            Sql &= "    + MAX(sales_tax) - ISNULL((SELECT SUM(pgl_amt)"
            Sql &= "    FROM PYMNTGLD WHERE pgl_ino=i.ivi_ino GROUP BY pgl_ino), 0)"
            Sql &= "    AS [balance]"
            Sql &= "    FROM INVITEMS i"
            Sql &= "    INNER JOIN INVOICE inv ON [ivi_ino]=[inv_ino]"
            Sql &= "    WHERE [ivi_ino] IN (SELECT [inv_ino] FROM INVOICE"
            Sql &= "    WHERE refund=1 AND inv_paid=0 AND inv_no=@cst_no AND voided=0) "
            Sql &= "    GROUP BY [ivi_ino]"
            Return Me.Database.GetAll(Sql)
        End Get
    End Property

    Public ReadOnly Property TotalPayments() As Double
        Get
            ' All-time total from payment gl
            Dim Sql As String = "SELECT ISNULL(SUM(pgl_amt), 0) FROM PYMNTGLD"
            Sql &= " INNER JOIN PAYMENTS ON pmt_recno=pgl_recno"
            Sql &= " WHERE pmt_no=" & Me.Database.Escape(Me.CustomerNo)
            Dim Amt As Double = Me.Database.GetOne(Sql)
            If Me.Database.LastQuery.Successful Then
                Return Amt
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property TotalInvoiced() As Double
        Get
            ' All-time total from invoice gl
            Dim Sql As String = "SELECT ISNULL(SUM(amount), 0) FROM invoice_gl"
            Sql &= " INNER JOIN INVOICE ON invoice_no=inv_ino"
            Sql &= " WHERE refund=0 AND voided=0 AND inv_no=" & Me.Database.Escape(Me.CustomerNo)
            Dim Amt As Double = Me.Database.GetOne(Sql)
            If Me.Database.LastQuery.Successful Then
                Return Amt
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property TotalRefunds() As Double
        Get
            ' All-time total from invoice gl
            Dim Sql As String = "SELECT ISNULL(SUM(amount), 0) FROM invoice_gl"
            Sql &= " INNER JOIN INVOICE ON invoice_no=inv_ino"
            Sql &= " WHERE refund=1 AND voided=0 AND inv_no=" & Me.Database.Escape(Me.CustomerNo)
            Dim Amt As Double = Me.Database.GetOne(Sql)
            If Me.Database.LastQuery.Successful Then
                Return Amt
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property AccountBalance() As Double
        Get
            Return Me.OutstandingInvoiceTotal - Me.TotalCredits
        End Get
    End Property

    Public ReadOnly Property TotalCredits() As Double
        Get
            Return Me.OutstandingRefundsTotal + Me.OverpaymentsTotal
        End Get
    End Property

    Public ReadOnly Property OverpaymentsTotal() As Double
        Get
            ' All-time total of overpayments account in payment gl
            ' (only gets unapplied since when applied to invoice would be canceled with a negative here)
            Dim GL As String = Me.Settings.GetValue("Overpayments GL", "")
            If GL.Length > 0 Then
                Dim Sql As String = "SELECT ISNULL(SUM(pgl_amt), 0) FROM PYMNTGLD"
                Sql &= " INNER JOIN PAYMENTS ON pmt_recno=pgl_recno"
                Sql &= " WHERE pmt_no=" & Me.Database.Escape(Me.CustomerNo)
                Sql &= " AND pgl_acc=" & Me.Database.Escape(GL)
                Dim Amt As Double = Me.Database.GetOne(Sql)
                If Me.Database.LastQuery.Successful Then
                    Return Amt
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property OverpaymentsReport() As DataTable
        Get
            Dim GL As String = Me.Settings.GetValue("Overpayments GL", "")
            If GL.Length > 0 Then
                Return Me.Database.GetAll("SELECT pgl_amt AS overpayment_amount, pmt_recno AS payment_recno," & _
                    " pmt_amt AS payment_total, pmt_date AS payment_date" & _
                    " FROM PYMNTGLD" & _
                    " INNER JOIN PAYMENTS ON pmt_recno=pgl_recno" & _
                    " WHERE pmt_no='" & Me.CustomerNo & "' AND pgl_acc=" & Me.Database.Escape(GL))
            Else
                Return New DataTable
            End If
        End Get
    End Property

    Public Function IsDuplicateName(ByVal Name As String) As Boolean
        Dim Sql As String = "SELECT * FROM ADDRESS WHERE cst_name=" & Me.Database.Escape(Name)
        If Me.CustomerNo.Length > 0 Then
            Sql &= " AND cst_no <> " & Me.Database.Escape(Me.CustomerNo)
        End If
        Dim Table As DataTable = Me.Database.GetAll(Sql)
        If Table.Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function


    Public Sub New(ByRef db As MyCore.Data.EasySql)
        Me.Database = db
        Me.Settings = New cSettings(Me.Database)
        Me.PopulateOffice()
        Me.PopulateCountries()
        Me.PopulateRateTables()
        Me.PopulateSalesTerritories()
        Me.PopulateServiceAreas()
        Me.PopulateStates()
        Me.PopulateTaxCodes()
        Me.PopulateTerms()
        Me.PopulateIndustries()
    End Sub

    Public Sub OpenAsNewVendor(Optional ByVal Clone As cCompany = Nothing)
        Me.Type = 3
        Me._IsVendor = True
        If Clone IsNot Nothing Then
            Me.Name = Clone.Name
            Me.Address1 = Clone.Address1
            Me.Address2 = Clone.Address2
            Me.City = Clone.City
            Me.State = Clone.State
            Me.County = Clone.County
            Me.Country = Clone.Country
            Me.Fax = Clone.Fax
            Me.Phone = Clone.Phone
            Me.Inactive = Clone.Inactive
            Me.Industry = Clone.Industry
            Me.ISO = Clone.ISO
            Me.OfficeNo = Clone.OfficeNo
            Me.OurCustomerNo = Clone.OurCustomerNo
            Me.Ownership = Clone.Ownership
            Me.ParentNo = Clone.ParentNo
            Me.TickerSymbol = Clone.TickerSymbol
            Me.WebSite = Clone.WebSite
            Me.Zip = Clone.Zip
        End If
        RaiseEvent Reload()
    End Sub

    Public Sub OpenAsNewCustomer(Optional ByVal Clone As cCompany = Nothing)
        Me.Type = 1
        Me._IsVendor = False
        If Clone IsNot Nothing Then
            Me.Name = Clone.Name
            Me.Address1 = Clone.Address1
            Me.Address2 = Clone.Address2
            Me.City = Clone.City
            Me.State = Clone.State
            Me.County = Clone.County
            Me.Country = Clone.Country
            Me.Fax = Clone.Fax
            Me.Phone = Clone.Phone
            Me.Inactive = Clone.Inactive
            Me.Industry = Clone.Industry
            Me.ISO = Clone.ISO
            Me.OfficeNo = Clone.OfficeNo
            Me.OurCustomerNo = Clone.OurCustomerNo
            Me.Ownership = Clone.Ownership
            Me.ParentNo = Clone.ParentNo
            Me.TickerSymbol = Clone.TickerSymbol
            Me.WebSite = Clone.WebSite
            Me.Zip = Clone.Zip
        End If
        RaiseEvent Reload()
    End Sub

    Public Sub Open(ByVal CustomerNo As String)
        Dim Sql As String = ""
        Sql &= "SELECT TOP 1"
        Sql &= " ADDRESS.*, template.name AS cert_template_name,"
        Sql &= " ct.vendor, terms.acct_ref AS terms_ref, tc.acct_ref AS tax_ref,"
        Sql &= " ISNULL(bil_name, '') AS billing_name, "
        Sql &= " ISNULL(bil_addr1, '') AS billing_address1, "
        Sql &= " ISNULL(bil_addr2, '') AS billing_address2,"
        Sql &= " ISNULL(bil_phone, '') AS billing_phone, "
        Sql &= " ISNULL(bil_city, '') AS billing_city, "
        Sql &= " ISNULL(bil_state, '') AS billing_state, "
        Sql &= " ISNULL(bil_zip, '') AS billing_zip,"
        Sql &= " ISNULL(bil_taxno, '') AS billing_taxno,"
        Sql &= " BILLADDR.fax AS billing_fax,"
        Sql &= " BILLADDR.country AS billing_country,"
        Sql &= " terms.name AS terms_name, terms.days_until_due, terms.days_until_late, "
        Sql &= Me.Database.DiffDays("(SELECT MAX(inv_date) FROM INVOICE WHERE inv_no=ADDRESS.cst_no AND voided=0)", Me.Database.Timestamp) & " AS [last_invoice],"
        Sql &= Me.Database.DiffDays("(SELECT MIN(inv_date) FROM INVOICE WHERE inv_no=ADDRESS.cst_no AND voided=0)", Me.Database.Timestamp) & " AS [first_invoice],"
        Sql &= Me.Database.DiffDays("(SELECT MIN(inv_date) FROM INVOICE WHERE inv_no=ADDRESS.cst_no AND inv_paid=0 AND voided=0)", Me.Database.Timestamp) & " AS [oldest_unpaid_invoice]"
        Sql &= " FROM ADDRESS"
        Sql &= " LEFT OUTER JOIN BILLADDR ON cst_no=bil_no"
        Sql &= " LEFT OUTER JOIN template ON template.id=cert_template_id"
        Sql &= " LEFT OUTER JOIN company_type ct ON ct.id=type"
        Sql &= " LEFT OUTER JOIN pay_status terms ON pay_status_id=terms.id"
        Sql &= " LEFT OUTER JOIN tax_code tc ON tax_code_id=tc.id"
        Sql &= " WHERE (cst_no = @customer_no)"
        Sql = Sql.Replace("@customer_no", Me.Database.Escape(CustomerNo))
        Dim r As DataRow = Me.Database.GetRow(Sql)
        If Not Me.Database.LastQuery.Successful Then
            Throw New Exception("Encountered a database error while trying to open company." & ControlChars.CrLf & ControlChars.CrLf & Me.Database.LastQuery.ErrorMsg & "--" & Me.Database.LastQuery.CommandText)
        ElseIf Me.Database.LastQuery.RowsReturned <> 1 Then
            Throw New Exception("Company with that customer number was not found." & ControlChars.CrLf & ControlChars.CrLf & Me.Database.LastQuery.CommandText)
        Else
            Me._CustomerNo = CustomerNo
            Me.AccountNo = Me.IsNull(r.Item("cst_acc"))
            Me.Address1 = Me.IsNull(r.Item("cst_addr1"))
            Me.Address2 = Me.IsNull(r.Item("cst_addr2"))
            Me.Alert = Me.IsNull(r.Item("cst_alert"), False)
            Me.AlertText = Me.IsNull(r.Item("cst_alerttext"))
            Me.BillingAddress1 = Me.IsNull(r.Item("billing_address1"))
            Me.BillingAddress2 = Me.IsNull(r.Item("billing_address2"))
            Me.BillingCity = Me.IsNull(r.Item("billing_city"))
            Me.BillingCountry = Me.IsNull(r.Item("billing_country"))
            Me.BillingFax = Me.IsNull(r.Item("billing_fax"))
            Me.BillingName = Me.IsNull(r.Item("billing_name"))
            Me.BillingPhone = Me.IsNull(r.Item("billing_phone"))
            Me.BillingState = Me.IsNull(r.Item("billing_state"))
            Me.BillingZip = Me.IsNull(r.Item("billing_zip"))
            Me.CertTemplate = r.Item("cert_template_id")
            Me.CertTemplateName = Me.IsNull(r.Item("cert_template_name"))
            Me.City = Me.IsNull(r.Item("cst_city"))
            Me.Country = Me.IsNull(r.Item("country"))
            Me.County = Me.IsNull(r.Item("cst_cnty"))
            Me.Fax = Me.IsNull(r.Item("cst_fax"))
            Me.GenerateCerts = Me.IsNull(r.Item("generate_certs"), False)
            Me.Industry = Me.IsNull(r.Item("industry_id"), 0)
            Me.ISO = Me.IsNull(r.Item("iso_customer"), False)
            Me.MarketingExclude = Me.IsNull(r.Item("marketing_exclude"), False)
            Me.Name = Me.IsNull(r.Item("cst_name"))
            Me.Notes = Me.IsNull(r.Item("notes"))
            Me.OfficeNo = Me.IsNull(r.Item("office"), 0)
            Me.OurCustomerNo = Me.IsNull(r.Item("cst_ourno"))
            Me.Ownership = Me.IsNull(r.Item("ownership"))
            Me.ParentNo = Me.IsNull(r.Item("parent_id"), 0)
            Me.Password = Me.IsNull(r.Item("password"))
            Me.Phone = Me.IsNull(r.Item("cst_phone"))
            Me.RateTable = Me.IsNull(r.Item("rate_table_id"), 0)
            Me.RoundtripMiles = r.Item("roundtrip_miles")
            Me.SalesTerritory = Me.IsNull(r.Item("sales_territory_id"), 0)
            Me.ServiceArea = Me.IsNull(r.Item("service_area_id"), 0)
            Me.ServiceNote = Me.IsNull(r.Item("service_note"))
            Me.State = Me.IsNull(r.Item("cst_state"))
            Me.TaxNo = Me.IsNull(r.Item("billing_taxno"))
            Me.TaxCode = Me.IsNull(r.Item("tax_code_id"), 0)
            Me.TaxExemptThrough = Me.IsNull(r.Item("tax_exempt_thru"), Nothing)
            Me.Terms = Me.IsNull(r.Item("pay_status_id"), 0)
            Me._TermsName = Me.IsNull(r.Item("terms_name"))
            Me._DaysToLate = Me.IsNull(r.Item("days_until_late"), 31)
            Me._DaysToDue = Me.IsNull(r.Item("days_until_due"), 30)
            Me.TickerSymbol = Me.IsNull(r.Item("stock_ticker"))
            Me.Type = Me.IsNull(r.Item("type"), 0)
            Me.WebSite = Me.IsNull(r.Item("cst_url"))
            Me.Zip = Me.IsNull(r.Item("cst_zip"))
            Me._IsVendor = r.Item("vendor")
            Me._DaysSinceFirstInvoice = Me.IsNull(r.Item("first_invoice"), -1)
            Me._DaysSinceLastInvoice = Me.IsNull(r.Item("last_invoice"), -1)
            Me._DaysSinceOldestUnpaidInvoice = Me.IsNull(r.Item("oldest_unpaid_invoice"), -1)
            Me.ServiceNote = Me.IsNull(r.Item("service_note"))
            Me.CustomText1 = Me.IsNull(r.Item("custom_text1"))
            Me.CustomText2 = Me.IsNull(r.Item("custom_text2"))
            Me.CustomText3 = Me.IsNull(r.Item("custom_text3"))
            Me.InvoiceLineItemSetting = r.Item("invoice_line_item_setting")
            Me.PrintInvoiceSetting = r.Item("print_invoice_setting")
            Me.PreferredDeliveryMethod = r.Item("delivery_method")
            Me.APEmailAddress = r.Item("ap_email_address")
            Me.Inactive = r.Item("inactive")
            Me.BlanketPO = r.Item("blanket_po")
            Me.AcctRef = Me.IsNull(r.Item("acct_ref"), Nothing)
            Me.AcctUpload = Me.IsNull(r.Item("acct_upload"), Nothing)
            Me.TermsRef = Me.IsNull(r.Item("terms_ref"), Nothing)
            Me.SalesTaxRef = Me.IsNull(r.Item("tax_ref"), Nothing)
            RaiseEvent Reload()
        End If
    End Sub

    Private Function IsNull(ByVal Value As Object, Optional ByVal DefaultValue As Object = "") As Object
        If Value Is DBNull.Value Then
            Return DefaultValue
        Else
            Return Value
        End If
    End Function

    Private Sub PopulateCountries()
        Me._Countries = Me.Database.GetAll("SELECT abbreviation, name FROM country ORDER BY sort, name")
    End Sub

    Private Sub PopulateStates()
        Me._States = Me.Database.GetAll("SELECT abbreviation FROM state ORDER BY sort, name")
    End Sub

    Private Sub PopulateTaxCodes()
        Me._TaxCode = Me.Database.GetAll("SELECT id, name FROM tax_code ORDER BY name")
    End Sub

    Private Sub PopulateRateTables()
        Me._RateTable = Me.Database.GetAll("SELECT id, name FROM rate_table")
    End Sub

    Private Sub PopulateOffice()
        Me._Office = Me.Database.GetAll("SELECT id, number, name FROM office ORDER BY sort, name")
    End Sub

    Private Sub PopulateTerms()
        Me._Terms = Me.Database.GetAll("SELECT id, abbreviation, name FROM pay_status ORDER BY sort, name")
    End Sub

    Public Function CompanyTypes() As DataTable
        Dim Sql As String = "SELECT id, abbreviation, name"
        Sql &= " FROM company_type"
        If Me.IsVendor Then
            Sql &= " WHERE vendor=1"
        Else
            Sql &= " WHERE vendor=0"
        End If
        Sql &= " ORDER BY sort, name"
        Return Me.Database.GetAll(Sql)
    End Function

    Private Sub PopulateSalesTerritories()
        Me._SalesTerritories = Me.Database.GetAll("SELECT id, abbreviation, name FROM sales_territory ORDER BY sort, name")
    End Sub

    Private Sub PopulateServiceAreas()
        Me._ServiceAreas = Me.Database.GetAll("SELECT id, name, technician FROM service_area ORDER BY name")
    End Sub

    Private Sub PopulateIndustries()
        Me._Industries = Me.Database.GetAll("SELECT id, name FROM industry ORDER BY sort, name")
    End Sub

    Public Sub Save()
        ' Check if complete
        If Not Me.IsComplete Then
            Throw New Exception("Not all required data has been supplied.")
        End If
        ' Continue
        Dim NewCustomerNo As String = Nothing
        Dim Params As New Collection
        ' If new
        If Me._CustomerNo.Length = 0 Then
            NewCustomerNo = Me.GetNextNumber
            Dim Sql As String = "INSERT INTO ADDRESS"
            Sql &= " (cst_no, cst_name, cst_addr1, cst_addr2, cst_city, cst_state, cst_zip,"
            Sql &= " cst_cnty, cst_phone, cst_fax, cst_url, type, country,"
            Sql &= " office, sales_territory_id, service_area_id, pay_status_id, rate_table_id, parent_id,"
            Sql &= " cst_chngt, tax_code_id, blanket_po)"
            Sql &= " VALUES ("
            Sql &= " @customer_no, @name, @address1, @address2, @city,@state, @zip, @county, @phone, @fax, @url, @type,"
            Sql &= " @country, @office, @territory, @service_area_id, @pay_status,"
            Sql &= " @rate_table_id, @parent_no, " & Me.Database.Timestamp & ", @tax_code_id, @blanket_po"
            Sql &= " )"
            Sql = Sql.Replace("@customer_no", Me.Database.Escape(NewCustomerNo))
            Sql = Sql.Replace("@name", Me.Database.Escape(Me.Name))
            Sql = Sql.Replace("@address1", Me.Database.Escape(Me.Address1))
            Sql = Sql.Replace("@address2", Me.Database.Escape(Me.Address2))
            Sql = Sql.Replace("@city", Me.Database.Escape(Me.City))
            Sql = Sql.Replace("@state", Me.Database.Escape(Me.State))
            Sql = Sql.Replace("@zip", Me.Database.Escape(Me.Zip))
            Sql = Sql.Replace("@county", Me.Database.Escape(Me.County))
            Sql = Sql.Replace("@country", Me.Database.Escape(Me.Country))
            Sql = Sql.Replace("@phone", Me.Database.Escape(Me.Phone))
            Sql = Sql.Replace("@fax", Me.Database.Escape(Me.Fax))
            Sql = Sql.Replace("@url", Me.Database.Escape(Me.WebSite))
            Sql = Sql.Replace("@type", Me.Database.Escape(Me.Type))
            Sql = Sql.Replace("@parent_no", Me.Database.Escape(Me.ParentNo))
            Sql = Sql.Replace("@tax_code_id", Me.Database.Escape(Me.TaxCode))
            Sql = Sql.Replace("@pay_status", Me.Database.Escape(Me.Terms))
            Sql = Sql.Replace("@rate_table_id", Me.Database.Escape(Me.RateTable))
            Sql = Sql.Replace("@office", Me.Database.Escape(Me.OfficeNo))
            Sql = Sql.Replace("@territory", Me.Database.Escape(Me.SalesTerritory))
            Sql = Sql.Replace("@service_area_id", Me.Database.Escape(Me.ServiceArea))
            Sql = Sql.Replace("@blanket_po", Me.Database.Escape(Me.BlanketPO))
            Me.Database.Execute(Sql)
        Else
            Dim Sql As String = ""
            Sql &= "UPDATE [ADDRESS]"
            Sql &= " SET "
            Sql &= " cst_name=@name,"
            Sql &= " cst_addr1=@address1,"
            Sql &= " cst_addr2=@address2,"
            Sql &= " cst_city=@city,"
            Sql &= " cst_state=@state,"
            Sql &= " cst_zip=@zip,"
            Sql &= " cst_cnty=@county,"
            Sql &= " country=@country,"
            Sql &= " type=@type,"
            Sql &= " parent_id=@parent_id,"
            Sql &= " cst_phone=@phone,"
            Sql &= " cst_fax=@fax,"
            Sql &= " cst_url=@homepage,"
            Sql &= " office=@office,"
            Sql &= " cst_off=@office,"
            Sql &= " sales_territory_id=@sales_territory_id,"
            Sql &= " cst_ourno=@our_no,"
            Sql &= " stock_ticker=@ticker,"
            Sql &= " industry_id=@industry_id,"
            Sql &= " cst_user=@last_updated_by,"
            Sql &= " cst_chngt=" & Me.Database.Timestamp & ","
            Sql &= " service_area_id=@service_area_id,"
            Sql &= " pay_status_id=@pay_status_id,"
            Sql &= " notes=@notes,"
            Sql &= " rate_table_id=@rate_table_id,"
            Sql &= " tax_exempt_thru=@tax_exempt_thru,"
            Sql &= " ownership=@ownership,"
            Sql &= " tax_code_id=@tax_code_id,"
            Sql &= " generate_certs=@generate_certs,"
            Sql &= " iso_customer=@iso_customer,"
            Sql &= " cert_template_id=@cert_template_id,"
            Sql &= " marketing_exclude=@marketing_exclude,"
            Sql &= " [password]=@password,"
            Sql &= " cst_alert=@alert,"
            Sql &= " cst_alerttext=@alert_text,"
            Sql &= " custom_text1=@custom_text1,"
            Sql &= " custom_text2=@custom_text2,"
            Sql &= " custom_text3=@custom_text3,"
            Sql &= " service_note=@service_note,"
            Sql &= " roundtrip_miles=@mileage,"
            Sql &= " accounting_customer_no=@accounting_customer_no,"
            Sql &= " print_invoice_setting=@print_invoice_setting,"
            Sql &= " invoice_line_item_setting=@invoice_line_item_setting,"
            Sql &= " inactive=@inactive,"
            Sql &= " delivery_method=@delivery_method,"
            Sql &= " ap_email_address=@ap_email_address,"
            Sql &= " blanket_po=@blanket_po,"
            Sql &= " acct_ref=@acct_ref,"
            Sql &= " acct_upload=@acct_upload"
            Sql &= " WHERE cst_no=@customer_no"
            Sql = Sql.Replace("@customer_no", Me.Database.Escape(Me.CustomerNo))
            Sql = Sql.Replace("@name", Me.Database.Escape(Me.Name))
            Sql = Sql.Replace("@address1", Me.Database.Escape(Me.Address1))
            Sql = Sql.Replace("@address2", Me.Database.Escape(Me.Address2))
            Sql = Sql.Replace("@city", Me.Database.Escape(Me.City))
            Sql = Sql.Replace("@state", Me.Database.Escape(Me.State))
            Sql = Sql.Replace("@zip", Me.Database.Escape(Me.Zip))
            Sql = Sql.Replace("@country", Me.Database.Escape(IIf(Me.Country Is Nothing, "", Me.Country)))
            Sql = Sql.Replace("@county", Me.Database.Escape(Me.County))
            Sql = Sql.Replace("@type", Me.Database.Escape(Me.Type))
            Sql = Sql.Replace("@parent_id", Me.Database.Escape(Me.ParentNo))
            Sql = Sql.Replace("@phone", Me.Database.Escape(Me.Phone))
            Sql = Sql.Replace("@fax", Me.Database.Escape(Me.Fax))
            Sql = Sql.Replace("@homepage", Me.Database.Escape(Me.WebSite))
            Sql = Sql.Replace("@office", Me.Database.Escape(Me.OfficeNo))
            Sql = Sql.Replace("@sales_territory_id", Me.Database.Escape(Me.SalesTerritory))
            Sql = Sql.Replace("@our_no", Me.Database.Escape(Me.OurCustomerNo))
            Sql = Sql.Replace("@ticker", Me.Database.Escape(Me.TickerSymbol))
            Sql = Sql.Replace("@industry_id", Me.Database.Escape(Me.Industry))
            Sql = Sql.Replace("@last_updated_by", Me.Database.Escape(Me.LastUpdatedBy))
            Sql = Sql.Replace("@service_area_id", Me.Database.Escape(Me.ServiceArea))
            Sql = Sql.Replace("@pay_status_id", Me.Database.Escape(Me.Terms))
            Sql = Sql.Replace("@notes", Me.Database.Escape(Me.Notes))
            Sql = Sql.Replace("@account_no", Me.Database.Escape(Me.AccountNo))
            Sql = Sql.Replace("@rate_table_id", Me.Database.Escape(Me.RateTable))
            Sql = Sql.Replace("@tax_exempt_thru", Me.Database.Escape(IIf(Me.TaxExemptThrough = Nothing, DBNull.Value, Me.TaxExemptThrough)))
            Sql = Sql.Replace("@ownership", Me.Database.Escape(Me.Ownership))
            Sql = Sql.Replace("@tax_code_id", Me.Database.Escape(Me.TaxCode))
            Sql = Sql.Replace("@generate_certs", Me.Database.Escape(Me.GenerateCerts))
            Sql = Sql.Replace("@iso_customer", Me.Database.Escape(Me.ISO))
            Sql = Sql.Replace("@cert_template_id", Me.Database.Escape(Me.CertTemplate))
            Sql = Sql.Replace("@password", Me.Database.Escape(Me.Password))
            Sql = Sql.Replace("@marketing_exclude", Me.Database.Escape(Me.MarketingExclude))
            Sql = Sql.Replace("@alert_text", Me.Database.Escape(Me.AlertText))
            Sql = Sql.Replace("@alert", Me.Database.Escape(Me.Alert))
            Sql = Sql.Replace("@service_note", Me.Database.Escape(Me.ServiceNote))
            Sql = Sql.Replace("@mileage", Me.Database.Escape(Me.RoundtripMiles))
            Sql = Sql.Replace("@custom_text1", Me.Database.Escape(Me.CustomText1))
            Sql = Sql.Replace("@custom_text2", Me.Database.Escape(Me.CustomText2))
            Sql = Sql.Replace("@custom_text3", Me.Database.Escape(Me.CustomText3))
            Sql = Sql.Replace("@accounting_customer_no", Me.Database.Escape(Me.AccountingCustomerNo))
            Sql = Sql.Replace("@invoice_line_item_setting", Me.Database.Escape(Me.InvoiceLineItemSetting))
            Sql = Sql.Replace("@print_invoice_setting", Me.Database.Escape(Me.PrintInvoiceSetting))
            Sql = Sql.Replace("@delivery_method", Me.Database.Escape(Me.PreferredDeliveryMethod))
            Sql = Sql.Replace("@ap_email_address", Me.Database.Escape(Me.APEmailAddress))
            Sql = Sql.Replace("@inactive", Me.Database.Escape(Me.Inactive))
            Sql = Sql.Replace("@blanket_po", Me.Database.Escape(Me.BlanketPO))
            Sql = Sql.Replace("@acct_ref", Me.Database.Escape(Me.AcctRef))
            Sql = Sql.Replace("@acct_upload", Me.Database.Escape(Me.AcctUpload))
            Me.Database.Execute(Sql)
            If Me.Database.LastQuery.Successful Then
                Sql = "UPDATE BILLADDR"
                Sql &= " SET"
                Sql &= " bil_name=@billing_name,"
                Sql &= " bil_addr1=@billing_address1,"
                Sql &= " bil_addr2=@billing_address2,"
                Sql &= " bil_city=@billing_city,"
                Sql &= " bil_state=@billing_state,"
                Sql &= " bil_zip=@billing_zip,"
                Sql &= " bil_phone=@billing_phone,"
                Sql &= " bil_taxno=@billing_taxno,"
                Sql &= " bil_user=@last_updated_by,"
                Sql &= " bil_chngt=" & Me.Database.Timestamp & ","
                Sql &= " country=@billing_country,"
                Sql &= " fax=@billing_fax"
                Sql &= " WHERE bil_no=@customer_no"
                Sql = Sql.Replace("@billing_name", Me.Database.Escape(Me.BillingName))
                Sql = Sql.Replace("@billing_address1", Me.Database.Escape(Me.BillingAddress1))
                Sql = Sql.Replace("@billing_address2", Me.Database.Escape(Me.BillingAddress2))
                Sql = Sql.Replace("@billing_city", Me.Database.Escape(Me.BillingCity))
                Sql = Sql.Replace("@billing_state", Me.Database.Escape(Me.BillingState))
                Sql = Sql.Replace("@billing_zip", Me.Database.Escape(Me.BillingZip))
                Sql = Sql.Replace("@billing_fax", Me.Database.Escape(Me.BillingFax))
                Sql = Sql.Replace("@billing_phone", Me.Database.Escape(Me.BillingPhone))
                Sql = Sql.Replace("@billing_country", Me.Database.Escape(Me.BillingCountry))
                Sql = Sql.Replace("@billing_taxno", Me.Database.Escape(Me.TaxNo))
                Sql = Sql.Replace("@last_updated_by", Me.Database.Escape(Me.LastUpdatedBy))
                Sql = Sql.Replace("@customer_no", Me.Database.Escape(Me.CustomerNo))
                Me.CreateBillingRecordIfNeeded()
                Me.Database.Execute(Sql)
            End If

        End If
        If Me.Database.LastQuery.Successful Then
            ' If new, set id
            If Me._CustomerNo = "" Then
                Me._CustomerNo = NewCustomerNo
                Me.IncrementNextNumber()
            End If
            RaiseEvent AfterSave(Me)
            ' Reopen
            Me.Open(Me._CustomerNo)
        Else
            Throw New Exception("Company not saved." & ControlChars.CrLf & ControlChars.CrLf & "Database error: " & Me.Database.LastQuery.ErrorMsg)
        End If
    End Sub

    Private Sub CreateBillingRecordIfNeeded()
        Dim Needed As Boolean = False
        If Me.BillingAddress1.Length > 0 Then Needed = True
        If Me.BillingAddress2.Length > 0 Then Needed = True
        If Me.BillingCity.Length > 0 Then Needed = True
        If Me.BillingFax.Length > 0 Then Needed = True
        If Me.BillingPhone.Length > 0 Then Needed = True
        If Me.BillingName.Length > 0 Then Needed = True
        If Me.BillingZip.Length > 0 Then Needed = True
        If Me.TaxNo.Length > 0 Then Needed = True
        If Needed Then
            Dim num As Integer = Me.Database.GetOne("SELECT COUNT(bil_id) FROM BILLADDR WHERE bil_no=" & Me.Database.Escape(Me.CustomerNo))
            If num = 0 Then
                Dim Sql As String = "INSERT INTO BILLADDR"
                Sql &= " (bil_no) VALUES ('" & Me.CustomerNo & "')"
                Me.Database.Execute(Sql)
            End If
        End If
    End Sub

    Private Function IsComplete() As Boolean
        If Me.Name.Trim.Length = 0 Then
            Return False
        End If
        If Me.Type = 0 Then
            Return False
        End If
        If Me.Type = 1 Or Me.Type = 2 Then
            If Me.TaxCode = 0 Then
                Return False
            End If
            If Me.RateTable = 0 Then
                Return False
            End If
            If Me.Terms = 0 Then
                Return False
            End If
            If Me.SalesTerritory = 0 Then
                Return False
            End If
            If Me.ServiceArea = 0 Then
                Return False
            End If
            If Me.OfficeNo = 0 Then
                Return False
            End If
        End If
        Return True
    End Function

    Public Function GetContacts() As DataTable
        Dim Sql As String = "SELECT *, ISNULL(cnt_last, '') + ', ' + ISNULL(cnt_first, '') AS display_name"
        Sql &= " FROM CONTACTS WHERE cnt_no= " & Me.Database.Escape(Me.CustomerNo)
        Sql &= " ORDER BY display_name"
        Dim Table As DataTable = Me.Database.GetAll(Sql)
        If Me.Database.LastQuery.Successful Then
            Return Table
        Else
            Dim Err As String = Me.Database.LastQuery.ErrorMsg
            Return Nothing
        End If
    End Function

    Public Function GetEquipment(Optional ByVal InactiveFitler As String = "all") As DataTable
        Dim Sql As String = "SELECT equip.*, item_category.name AS type,"
        Sql &= " name = CASE WHEN equip.dep_ser IS NOT NULL THEN equip.dep_manuf + ' ' + equip.dep_mod  + ' (' + equip.dep_ser + ')'"
        Sql &= " ELSE equip.dep_manuf + ' ' + equip.dep_mod END"
        Sql &= " FROM DEPREC AS equip INNER JOIN"
        Sql &= " ADDRESS a ON equip.dep_loc = a.cst_no LEFT OUTER JOIN"
        Sql &= " item_category ON equip.dep_type = item_category.id"
        Sql &= " WHERE a.cst_no=" & Me.Database.Escape(Me.CustomerNo)
        If InactiveFitler = "active" Then
            Sql &= " AND equip.inactive=0"
        ElseIf InactiveFitler = "inactive" Then
            Sql &= " AND equip.inactive=1"
        End If
        Sql &= " ORDER BY name"
        Dim TAble As DataTable = Me.Database.GetAll(Sql)
        If Me.Database.LastQuery.Successful Then
            Return TAble
        Else
            Throw New Exception(Me.Database.LastQuery.ErrorMsg)
        End If
    End Function

    Public Function GetServiceOrders() As DataTable
        Dim Sql As String = "SELECT so.id AS service_order_id, so.location_id, so.office, so.cal_agreement_id,"
        Sql &= " so.notes, so.date_due, "
        Sql &= " (SELECT MIN(date_start) FROM schedule WHERE reference_id=so.id AND deleted=0) AS date_scheduled,"
        Sql &= " so.date_completed, so.invoice_id, so.contract, so.techs, so.helpers, "
        Sql &= " so.return_trip_required, so.date_created, so.created_by, work_location='',"
        Sql &= " CASE "
        Sql &= " WHEN so.voided=1 THEN 'Voided'"
        Sql &= " WHEN so.invoice_id > 0 THEN 'Invoiced'"
        Sql &= " WHEN so.date_completed IS NOT NULL THEN 'Completed'"
        Sql &= " WHEN so.date_scheduled IS NOT NULL THEN 'Scheduled'"
        Sql &= " ELSE 'Received'"
        Sql &= " END AS stage,"
        Sql &= " (SELECT name FROM service_order_type sot WHERE sot.id=so.service_order_type) AS type"
        Sql &= " FROM service_order so"
        Sql &= " WHERE so.location_id=" & Me.Database.Escape(Me.CustomerNo)
        Sql &= " ORDER BY so.date_created DESC"
        Dim Table As DataTable = Me.Database.GetAll(Sql)
        If Me.Database.LastQuery.Successful Then
            Return Table
        Else
            Throw New Exception("Couldn't get service orders. Error from database: " & Me.Database.LastQuery.ErrorMsg)
            Return Nothing
        End If
    End Function

    Public Function GetSalesOrders() As DataTable
        Dim Sql As String = "SELECT so.*, "
        Sql &= " s.cst_name AS ship_to_name, s.cst_city + ', ' + s.cst_state AS ship_to_city,"
        Sql &= " b.cst_name AS bill_to_name, b.cst_city + ', ' + b.cst_state AS bill_to_city"
        Sql &= " FROM sales_order so"
        Sql &= " LEFT JOIN ADDRESS s ON so.ship_to=s.cst_no"
        Sql &= " LEFT JOIN ADDRESS b ON so.bill_to=b.cst_no"
        Sql &= " WHERE so.bill_to=" & Me.Database.Escape(Me.CustomerNo)
        Sql &= " OR so.ship_to=" & Me.Database.Escape(Me.CustomerNo)
        Sql &= " ORDER BY so.date_created DESC"
        Dim Table As DataTable = Me.Database.GetAll(Sql)
        If Me.Database.LastQuery.Successful Then
            Return Table
        Else
            Throw New Exception("Error getting sales orders. Error from database: " & Me.Database.LastQuery.ErrorMsg)
        End If
    End Function

    Public Function GetRentalOrders() As DataTable
        Dim Sql As String = "SELECT ro.*, "
        Sql &= " s.cst_name AS ship_to_name, s.cst_city + ', ' + s.cst_state AS ship_to_city,"
        Sql &= " b.cst_name AS bill_to_name, b.cst_city + ', ' + b.cst_state AS bill_to_city"
        Sql &= " FROM rental_order ro"
        Sql &= " LEFT JOIN ADDRESS s ON ro.ship_to_no=s.cst_no"
        Sql &= " LEFT JOIN ADDRESS b ON ro.bill_to_no=b.cst_no"
        Sql &= " WHERE ro.ship_to_no=" & Me.Database.Escape(Me.CustomerNo)
        Sql &= " OR ro.bill_to_no=" & Me.Database.Escape(Me.CustomerNo)
        Sql &= " ORDER BY ro.date_created DESC"
        Return Me.Database.GetAll(Sql)
    End Function

    Public Function GetPurchaseOrders(Optional ByVal IncludeRMAs As Boolean = False) As DataTable
        Dim Sql As String = "SELECT *,"
        Sql &= " CASE [rma] WHEN 0 THEN 'PO' ELSE 'RMA' END AS [type],"
        Sql &= " CASE WHEN (SELECT SUM(quantity) FROM purchase_order_item WHERE po_no=po.po_no AND received < quantity) > 0 THEN 'N' ELSE 'Y' END AS is_received"
        Sql &= " FROM purchase_order po WHERE vendor_no=@customer_no"
        If Not IncludeRMAs Then
            Sql &= " AND rma=0"
        End If
        Sql &= " ORDER BY po.po_date DESC"
        Sql = Sql.Replace("@customer_no", Me.Database.Escape(Me.CustomerNo))
        Dim Table As DataTable = Me.Database.GetAll(Sql)
        If Me.Database.LastQuery.Successful Then
            Return Table
        Else
            Throw New Exception("Error getting purchase orders.  Error from database: " & Me.Database.LastQuery.ErrorMsg)
        End If
    End Function

    Public Function GetInvoices(ByVal Start As DateTime, ByVal [End] As DateTime, Optional ByVal Paid As Integer = 0, Optional ByVal IncludeRefunds As Boolean = False) As DataTable
        Dim Sql As String = ""
        Sql &= "SELECT * FROM"
        Sql &= " ("
        Sql &= " SELECT [inv_id] AS [invoice_id], i.inv_ino AS [invoice_no], [inv_date] AS [date],"
        Sql &= " [sub_total], [tax], "
        Sql &= " ISNULL(inv_shch, 0) AS [shipping], [total], ISNULL([payments], 0) AS [payments], "
        Sql &= " (ISNULL([total], 0)-ISNULL([payments], 0)) AS [balance], refund,"
        Sql &= " CASE [inv_paid] WHEN 1 THEN 'Paid' ELSE 'Unpaid' END AS [status],"
        Sql &= " CASE [refund] WHEN 0 THEN 'Invoice' ELSE 'Refund' END AS [type]"
        Sql &= " FROM INVOICE AS i"
        Sql &= " LEFT JOIN ("
        Sql &= "    SELECT [inv_ino], ISNULL(SUM([ivi_qua] * [ivi_cost]), 0) AS [sub_total],"
        Sql &= "    MAX(sales_tax) AS [tax],"
        Sql &= "    (ISNULL(SUM([ivi_qua] * [ivi_cost]), 0) + ISNULL(AVG([inv_shch]), 0) + MAX(sales_tax)) AS [total]"
        Sql &= "    FROM INVOICE LEFT OUTER JOIN INVITEMS ON inv_ino=ivi_ino"
        Sql &= "    WHERE [inv_ino] IN (SELECT [inv_ino] FROM INVOICE WHERE [inv_no] = @customer_no)"
        Sql &= "    GROUP BY [inv_ino]) "
        Sql &= " AS d ON i.inv_ino=d.inv_ino"
        Sql &= " LEFT OUTER JOIN ("
        Sql &= "    SELECT SUM(pgl_amt) AS [payments], pgl_ino"
        Sql &= "    FROM PYMNTGLD "
        Sql &= "    WHERE pgl_ino IN (SELECT [inv_ino] FROM INVOICE WHERE [inv_no] = @customer_no)"
        Sql &= "    GROUP BY [pgl_ino]) "
        Sql &= " AS p ON i.inv_ino=p.pgl_ino"
        Sql &= " WHERE voided=0 AND ([inv_no] = @customer_no OR [ship_to_no] = @customer_no)"
        If Not IncludeRefunds Then
            Sql &= " AND refund=0"
        End If
        Sql &= " AND ([inv_date] BETWEEN @start_date AND @end_date) "
        Sql &= " ) AS [results]"
        Sql &= " WHERE [status] LIKE @status"
        Sql &= " ORDER BY [date] DESC"
        Select Case Paid
            Case -1
                Sql = Sql.Replace("@status", Me.Database.Escape("Unpaid"))
            Case 0
                Sql = Sql.Replace("@status", Me.Database.Escape("%"))
            Case 1
                Sql = Sql.Replace("@status", Me.Database.Escape("Paid"))
        End Select
        Sql = Sql.Replace("@customer_no", Me.Database.Escape(Me.CustomerNo))
        Sql = Sql.Replace("@start_date", Me.Database.Escape(Start))
        Sql = Sql.Replace("@end_date", Me.Database.Escape([End]))
        Dim Table As DataTable = Me.Database.GetAll(Sql)
        If Not Me.Database.LastQuery.Successful Then
            Throw New Exception(Me.Database.LastQuery.ErrorMsg & " --- " & Me.Database.LastQuery.CommandText)
        Else
            If IncludeRefunds Then
                For Each Row As DataRow In Table.Rows
                    If Row.Item("refund") = 1 Then
                        Row.Item("total") = Row.Item("total") * -1
                        Row.Item("shipping") = Row.Item("shipping") * -1
                        Row.Item("tax") = Row.Item("tax") * -1
                        Row.Item("sub_total") = Row.Item("sub_total") * -1
                        Row.Item("payments") = Row.Item("payments") * -1
                        Row.Item("balance") = Row.Item("balance") * -1
                    End If
                Next
            End If
            Return Table
        End If
    End Function

    Public Function GetInteractions() As DataTable
        Dim Sql As String = "SELECT journal.contact_name, journal.customer_no, journal.id AS journal_id, journal.entry_type_id, journal.subject, journal.memo,"
        Sql &= " journal.created_by, journal.created_date, journal.touch_date, journal.touch_by, journal_entry_type.name AS entry_type, "
        Sql &= " CASE journal.department WHEN 1 THEN 'Sales' WHEN 2 THEN 'Service' WHEN 3 THEN 'Collections' WHEN 4 THEN 'Rental' ELSE 'Purchase' END AS department_name,"
        Sql &= " CASE journal.initiator WHEN 1 THEN 'Us' ELSE 'Them' END AS initiator_name,"
        Sql &= " journal.department, journal.initiator"
        Sql &= " FROM journal"
        Sql &= " INNER JOIN journal_entry_type ON journal.entry_type_id = journal_entry_type.id "
        Sql &= " WHERE journal.customer_no=" & Me.Database.Escape(Me.CustomerNo)
        Sql &= " ORDER BY journal.touch_date DESC"
        Return Me.Database.GetAll(Sql)
    End Function

    Public Function GetCalAgreements(Optional ByVal StatusFilter As String = "All") As DataTable
        Dim Sql As String = "SELECT cal.*, ship.cst_name AS ship_to_name,"
        Sql &= " (ship.cst_city + ', ' + ship.cst_state) AS ship_to_city, "
        Sql &= " date_next,"
        Sql &= " CASE cal.frequency_months WHEN  0 THEN cal.frequency_days ELSE cal.frequency_months END AS frequency_text,"
        Sql &= " CASE WHEN cal.date_end IS NULL THEN 'No' WHEN cal.date_end > " & Me.Database.Timestamp & " THEN 'No' ELSE 'Yes' END AS needs_review,"
        Sql &= " CASE WHEN cal.canceled = 0 THEN 'Current' ELSE 'Canceled' END AS status,"
        Sql &= " ISNULL((SELECT COUNT(cae.equipment_id) FROM cal_agreement_equipment cae WHERE cae.agreement_id=cal.id GROUP BY cae.agreement_id), 0) AS number_of_pieces"
        Sql &= " FROM cal_agreement cal"
        Sql &= " LEFT OUTER JOIN ADDRESS ship ON cal.ship_to_no=ship.cst_no"
        ' For this company only
        Sql &= " WHERE cal.ship_to_no=@customer_no"
        ' Must not be a non recurring or canceled agreement
        If StatusFilter = "Current" Or StatusFilter = "Active" Then
            Sql &= " AND cal.canceled=0"
        ElseIf StatusFilter = "Canceled" Then
            Sql &= " AND cal.canceled=1"
        End If
        Sql &= " ORDER BY cal.date_created DESC"
        Sql = Sql.Replace("@customer_no", Me.Database.Escape(Me.CustomerNo))
        'Dim Err As New MyCore.Gravity.ErrorBox("blah", "blah", Sql)
        Dim Table As DataTable = Me.Database.GetAll(Sql)
        For Each Row As DataRow In Table.Rows
            If Row.Item("canceled") Then
                Row.Item("date_next") = DBNull.Value
            End If
        Next
        Return Table
    End Function

    Public Function GetPayments() As DataTable
        Dim Sql As String = "SELECT DISTINCT pmt_recno AS record_no, pmt_date AS payment_date, pmt_chck AS check_no,"
        Sql &= " amount = SUM(pgl_amt)"
        Sql &= " FROM PYMNTGLD"
        Sql &= " INNER JOIN PAYMENTS ON pgl_recno=pmt_recno"
        Sql &= " WHERE pmt_no=" & Me.Database.Escape(Me.CustomerNo)
        Sql &= " GROUP BY pmt_recno, pmt_date, pmt_chck"
        Sql &= " ORDER BY payment_date DESC"
        Return Me.Database.GetAll(Sql)
    End Function

    Public Function GetProjects(Optional ByVal SelectFields As String = "*") As DataTable
        Return Me.Database.GetAll("SELECT " & SelectFields & " FROM project WHERE customer_no=" & Me.Database.Escape(Me.CustomerNo))
    End Function

    Public Function GetCustomerStatement(Optional ByVal Id As Integer = 0) As MyCore.GravityDocument.gDocument
        Dim Statement As New cCustomerStatement(Me.Database)
        If Id > 0 Then
            Statement.TemplateId = Id
        End If
        Return Statement.ToGravityDocument(Me)
    End Function

    Public Function StatementTable() As DataTable()
        ' Dimension variables
        Dim DaysToCompound As Integer = 30
        Dim Interest As Double = 2
        Dim PerDay As Double = Interest / DaysToCompound
        ' Create Line Items Table
        Dim Table As New DataTable
        Table.Columns.Add("date")
        Table.Columns.Add("invoice_no")
        Table.Columns.Add("invoice_amount")
        Table.Columns.Add("check_no")
        Table.Columns.Add("payment")
        Table.Columns.Add("balance")
        ' Time Period Table
        Dim Due As New DataTable
        Due.Columns.Add("current")
        Due.Columns.Add("1to30")
        Due.Columns.Add("31to60")
        Due.Columns.Add("61to90")
        Due.Columns.Add("91plus")
        Due.Rows.Add(Due.NewRow)
        Due.Rows(0).Item("current") = 0
        Due.Rows(0).Item("1to30") = 0
        Due.Rows(0).Item("31to60") = 0
        Due.Rows(0).Item("61to90") = 0
        Due.Rows(0).Item("91plus") = 0
        ' Prepare variables
        Dim Balance As Double = 0
        Dim InvBalance As Double = 0
        Dim LateCharges As Double = 0
        ' Get Oustanding Invoices
        Dim Sql As String = ""
        Sql &= " SELECT inv_ino, inv_date, ISNULL(inv_lchp, 0) AS late_charge_percent,"
        Sql &= " ("
        Sql &= "     SELECT (ISNULL(SUM([ivi_qua] * [ivi_cost]), 0) + ISNULL(AVG([inv_shch]), 0) + MAX(sales_tax))"
        Sql &= "     FROM INVITEMS, INVOICE WHERE [ivi_ino] = inv.inv_ino AND [ivi_ino]=[inv_ino] GROUP BY [ivi_ino]"
        Sql &= " ) AS total"
        Sql &= " FROM INVOICE inv"
        Sql &= " WHERE inv_paid=0 AND voided=0"
        Sql &= " AND inv_no=" & Me.Database.Escape(Me.CustomerNo)
        Sql &= " ORDER BY inv_date"
        Dim Invoices As DataTable = Me.Database.GetAll(Sql)
        ' Loop through invoices
        Dim Inv As DataRow
        For Each Inv In Invoices.Rows
            ' Terms
            DaysToCompound = Me.DaysToDue + Me.DaysToLate - 1
            ' Add Balance
            Dim InvAmt As Double = IIf(Inv.Item("total") Is DBNull.Value, 0, Inv.Item("total"))
            Balance += InvAmt
            InvBalance = InvAmt
            LateCharges = 0
            ' Set late charge
            Interest = Inv.Item("late_charge_percent")
            PerDay = Interest / DaysToCompound
            ' Add Statement Row for Initial Invoice
            Dim r As DataRow = Table.NewRow
            r.Item("date") = Format(Inv.Item("inv_date"), "MM/dd/yy")
            r.Item("invoice_no") = Inv.Item("inv_ino")
            r.Item("invoice_amount") = Format(InvAmt, "c")
            r.Item("check_no") = ""
            r.Item("payment") = ""
            r.Item("balance") = Format(Balance, "c")
            Table.Rows.Add(r)
            ' Get Payments
            Sql = "SELECT DISTINCT pmt_recno AS record_no, pmt_date AS payment_date, pmt_chck AS check_no,"
            Sql &= " SUM(pgl_amt) AS amount"
            Sql &= " FROM PYMNTGLD"
            Sql &= " INNER JOIN PAYMENTS ON pgl_recno=pmt_recno"
            Sql &= " WHERE pgl_ino=" & Me.Database.Escape(Inv.Item("inv_ino"))
            Sql &= " GROUP BY pmt_recno, pmt_date, pmt_chck"
            Dim Payments As DataTable = Me.Database.GetAll(Sql)
            ' If payments have been made
            If Payments.Rows.Count > 0 Then
                Dim DaysSinceInvoice As Integer = DateDiff(DateInterval.Day, CType(Inv.Item("inv_date"), Date), Today)
                Dim DaysSincePrevious As Integer = 0
                Dim DatePrevious As Date = Inv.Item("inv_date")
                ' Loop through payments
                Dim Pmt As DataRow
                For Each Pmt In Payments.Rows
                    ' Calculate Late Charges
                    DaysSincePrevious = DateDiff(DateInterval.Day, DatePrevious, CType(Pmt.Item("payment_date"), Date))
                    If DaysSinceInvoice > DaysToCompound Then
                        If InvBalance > 0 Then
                            LateCharges += InvBalance * (PerDay / 100) * DaysSincePrevious
                        End If
                    End If
                    ' Subtract from balance
                    Balance -= IIf(Pmt.Item("amount") Is DBNull.Value, 0, Pmt.Item("amount"))
                    InvBalance -= IIf(Pmt.Item("amount") Is DBNull.Value, 0, Pmt.Item("amount"))
                    ' Add Statement Row for Payment
                    r = Table.NewRow
                    r.Item("date") = "  " & Format(Pmt.Item("payment_date"), "MM/dd/yy")
                    r.Item("invoice_no") = Inv.Item("inv_ino")
                    r.Item("invoice_amount") = ""
                    r.Item("check_no") = Pmt.Item("check_no")
                    r.Item("payment") = Format(IIf(Pmt.Item("amount") Is DBNull.Value, 0, Pmt.Item("amount")), "c")
                    r.Item("balance") = Format(Balance, "c")
                    Table.Rows.Add(r)
                    ' Set this as previous payment
                    DatePrevious = Pmt.Item("payment_date")
                Next
                ' Update late charges from date of previous payment until today
                DaysSincePrevious = DateDiff(DateInterval.Day, DatePrevious, Today)
                If DaysSinceInvoice > DaysToCompound Then
                    If InvBalance > 0 Then
                        LateCharges += InvBalance * (PerDay / 100) * DaysSincePrevious
                    End If
                End If
                ' Final late charges
                InvBalance += LateCharges
                Balance += LateCharges
            Else
                ' If we're more than 30 days out from the invoice
                Dim DaysSinceInvoice As Integer = DateDiff(DateInterval.Day, CType(Inv.Item("inv_date"), Date), Today)
                If DaysSinceInvoice > DaysToCompound Then
                    If InvBalance > 0 Then
                        LateCharges = InvBalance * (PerDay / 100) * DaysSinceInvoice
                    End If
                    InvBalance += LateCharges
                    Balance += LateCharges
                End If
            End If
            ' Add Late charges to statement
            If LateCharges > 0 Then
                ' Add Statement Row for Late Charges
                r = Table.NewRow
                r.Item("date") = ""
                r.Item("invoice_no") = "Late"
                r.Item("invoice_amount") = Format(LateCharges, "c")
                r.Item("check_no") = ""
                r.Item("payment") = ""
                r.Item("balance") = Format(Balance, "c")
                Table.Rows.Add(r)
            End If
            ' Calculate what time period
            Dim Diff As Integer = DateDiff(DateInterval.Day, CType(Inv.Item("inv_date"), Date), Today)
            If Diff <= 30 Then
                Due.Rows(0).Item("current") += InvBalance
            ElseIf Diff < 31 Then
                Due.Rows(0).Item("1to30") += InvBalance
            ElseIf Diff < 61 Then
                Due.Rows(0).Item("31to60") += InvBalance
            ElseIf Diff < 91 Then
                Due.Rows(0).Item("61to90") += InvBalance
            Else
                Due.Rows(0).Item("91plus") += InvBalance
            End If
        Next
        ' Overpayments
        Dim Overpd As Double = Me.OverpaymentsTotal
        If Overpd > 0 Then
            ' Subtract from balance
            Balance -= Overpd
            ' Add Row
            Dim r As DataRow = Table.NewRow
            r.Item("date") = ""
            r.Item("invoice_no") = "Overpayment"
            r.Item("invoice_amount") = Format(Overpd, "c")
            r.Item("check_no") = ""
            r.Item("payment") = ""
            r.Item("balance") = Format(Balance, "c")
            Table.Rows.Add(r)
            ' Subtract overpayment from current
            Due.Rows(0).Item("current") -= Overpd
        End If
        ' Return
        Dim ReturnVal(1) As DataTable
        ReturnVal(0) = Table
        ReturnVal(1) = Due
        Return ReturnVal
    End Function

    Public Sub IncrementNextNumber()
        Me.Database.Execute("UPDATE next_number SET number=number+1 WHERE name='company'")
    End Sub

    Public Function GetNextNumber() As Integer
        Return Me.Database.GetOne("SELECT number FROM next_number WHERE name='company'")
    End Function

End Class
