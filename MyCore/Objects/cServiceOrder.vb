Imports MyCore.Data

Public Class cServiceOrder

    Dim _Id As Integer

    Public ProjectId As Integer = 0
    Public BillToNo As String = ""
    Public ShipToNo As String = ""
    Public CalledInBy As String = ""
    Public CalledInByID As Integer = 0
    Public ContactName As String = ""
    Public ContactID As Integer = 0
    Public Office As Integer = 0
    Public Shop As Boolean = False
    Public TakenBy As String = ""
    Public Po As String = ""
    Public Notes As String = ""
    Public Truck As Integer = 1
    Public RateType As Integer = 2
    Public WorkType As Integer = 1
    Public ServiceOrderType As Integer = 1
    Public PickUp As Boolean = False
    Public Contract As Boolean = False
    Public RateTable As Integer = 0
    Public DateDue As Date = Nothing
    Public CalAgreement As Integer = 0
    Public FollowUp As Integer = 0
    Public TaxCode As Integer = 1
    Public EstimatedHours As Double = 0
    Public Techs As Integer = 1
    Public Helpers As Integer = 0
    Public AssignedTo As String = ""
    Public InternalNotes As String = ""
    Public DateCompleted As Date = Nothing
    Public ApprovedBy As String = ""
    Public ApprovedByID As Integer = 0
    Public ChargeTo As ChargeToType = ChargeToType.Invoice
    Public InvoiceNo As Integer = 0
    Public Freight As Double = 0
    Public JobSummary As String = ""
    Public DateCreated As DateTime = Now
    Public CreatedBy As String = ""
    Public DateLastUpdated As DateTime = Now
    Public LastUpdateBy As String = ""
    Public DownloadToPrimaryTechOnly As Boolean = True
    Public StationId As Integer = 0
    Public PriorityID As Integer = 0
    Public TechSignature As String = ""
    Public Voided As Boolean = False
    Public AgreementAdvancedNextDue As Integer = 0

    Public ShipTo As Address
    Public BillTo As Address

    Dim _ShipToName As String = ""

    Public WorkOrders As DataTable
    Public LineItems As DataTable

    Public WorkTypes As DataTable
    Public RateTables As DataTable
    Public TruckTypes As DataTable
    Public Offices As DataTable
    Public TaxCodes As DataTable
    Public ServiceOrderTypes As DataTable

    Public Event Saved(ByVal ServiceOrder As cServiceOrder)

    Structure Address
        Dim Name As String
        Dim Address1 As String
        Dim Address2 As String
        Dim City As String
        Dim State As String
        Dim Zip As String
        Dim Fax As String
        Dim Phone As String
        Dim Terms As String
    End Structure

    Enum ChargeToType
        Invoice = 0
        SalesOrder = 1
        RentalOrder = 2
        ServiceOrder = 3
        NoCharge = 4
    End Enum

    Public Database As MyCore.Data.EasySql
    Public Event Reload()
    Public Event Progress(ByVal Percent As Integer, ByVal Action As String)

    Public ReadOnly Property OrderNo() As Integer
        Get
            Return Me._Id
        End Get
    End Property

    Public ReadOnly Property ShipToName() As String
        Get
            Return Me._ShipToName
        End Get
    End Property

    Public ReadOnly Property Completed() As Boolean
        Get
            If Me.DateCompleted = Nothing Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    Public ReadOnly Property TravelTotal() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                If Row.Item("item_type_id") IsNot DBNull.Value Then
                    If Row.Item("item_type_id") = 3 Then
                        Amt += Row.Item("quantity") * Row.Item("price")
                    End If
                End If
            Next
            Return Amt
        End Get
    End Property

    Public ReadOnly Property LaborTotal() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                If Row.Item("item_type_id") IsNot DBNull.Value Then
                    If Row.Item("item_type_id") = 2 Then
                        Amt += Row.Item("quantity") * Row.Item("price")
                    End If
                End If
            Next
            Return Amt
        End Get
    End Property

    Public ReadOnly Property PartsTotal() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                Try
                    If Row.Item("item_type_id") = 1 Then
                        Amt += Row.Item("quantity") * Row.Item("price")
                    End If
                Catch ex As Exception
                    ' Ignore
                End Try
            Next
            Return Amt
        End Get
    End Property

    Public ReadOnly Property SurchargesTotal() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                If Row.Item("item_type_id") IsNot DBNull.Value Then
                    If Row.Item("item_type_id") = 5 Then
                        Amt += Math.Round(Row.Item("quantity") * Row.Item("price"), 2)
                    End If
                End If
            Next
            Return Amt
        End Get
    End Property

    Public ReadOnly Property DiscountsTotal() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                If Row.Item("item_type_id") IsNot DBNull.Value Then
                    If Row.Item("item_type_id") = 6 Then
                        Amt += Math.Round(Row.Item("quantity") * Row.Item("price"), 2)
                    End If
                End If
            Next
            Return Amt
        End Get
    End Property

    Public ReadOnly Property OtherTaxTotal() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                If Row.Item("item_type_id") IsNot DBNull.Value Then
                    If Row.Item("item_type_id") = 7 Then
                        Amt += Math.Round(Row.Item("quantity") * Row.Item("price"), 2)
                    End If
                End If
            Next
            Return Amt
        End Get
    End Property

    Public ReadOnly Property SpecialTotal() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                If Row.Item("item_type_id") IsNot DBNull.Value Then
                    If Row.Item("item_type_id") = 8 Then
                        Amt += Math.Round(Row.Item("quantity") * Row.Item("price"), 2)
                    End If
                End If
            Next
            Return Amt
        End Get
    End Property

    Public ReadOnly Property LineItemsTotal() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                Dim qty As Double = IIf(Row.Item("quantity") Is DBNull.Value, 1, Row.Item("quantity"))
                Dim price As Double = IIf(Row.Item("price") Is DBNull.Value, 0, Row.Item("price"))
                Amt += qty * price
            Next
            Return Amt
        End Get
    End Property

    Public ReadOnly Property WorkOrderItemsTotal() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.WorkOrderItems.Rows
                Amt += Row.Item("total_price")
            Next
            Return Amt
        End Get
    End Property

    Public ReadOnly Property WorkOrderPartsTotal() As Double
        Get
            ' Parts
            Dim WoParts As Double
            Try
                WoParts = Me.Database.GetOne("SELECT SUM(quantity*unit_price) FROM work_order_item" & _
                    " WHERE work_order_id IN (" & _
                    "SELECT id FROM work_order WHERE service_order_id=" & Me.OrderNo & _
                    ") AND part_no <> '' AND part_no IS NOT NULL")
            Catch ex As InvalidCastException
                WoParts = 0
            End Try
            Return WoParts
        End Get
    End Property

    Public ReadOnly Property WorkOrderAdditionalTotal() As Double
        Get
            ' Additional Charges
            Dim WoAdditional As Double
            Try
                WoAdditional = Me.Database.GetOne("SELECT SUM(quantity*unit_price) FROM work_order_item" & _
                    " WHERE work_order_id IN (" & _
                    "SELECT id FROM work_order WHERE service_order_id=" & Me.OrderNo & _
                    ") AND (part_no = '' OR part_no IS NULL)")
            Catch ex As InvalidCastException
                WoAdditional = 0
            End Try
            Return WoAdditional
        End Get
    End Property

    Public ReadOnly Property TaxStatuses() As DataTable
        Get
            Return Me.Database.GetAll("SELECT id, code, description, taxable FROM tax_status ORDER BY code")
        End Get
    End Property

    Public ReadOnly Property Priorities() As DataTable
        Get
            Return Me.Database.GetAll("SELECT id, name, color FROM priority ORDER BY sort, name")
        End Get
    End Property

    Public ReadOnly Property DateScheduled() As Date
        Get

        End Get
    End Property

    Public ReadOnly Property SalesOrderID() As Integer
        Get
            Return Me.Database.GetOne("SELECT id FROM sales_order WHERE service_order_id=" & Me.OrderNo)
        End Get
    End Property

    Public ReadOnly Property Forms() As DataTable
        Get
            Dim sql As String
            sql = "SELECT f.id, f.name, date_created, created_by, form_type_id,"
            sql &= " ft.name AS form_type, f.protected"
            sql &= " FROM [form] f LEFT JOIN form_type ft ON f.form_type_id=ft.id"
            sql &= " WHERE reference_id=" & Me._Id & " AND (form_type_id=1 OR form_type_id=4)"
            sql &= " ORDER BY date_created"
            Return Me.Database.GetAll(sql)
        End Get
    End Property

    Public ReadOnly Property LineItemsForSalesTaxCalculation() As cTaxCode.LineItem()
        Get
            Dim WOItems As DataTable = Me.WorkOrderItems
            Dim num As Integer = 0
            If Me.LineItems.Rows.Count > 0 Then
                num += Me.LineItems.Rows.Count
            End If
            If WOItems.Rows.Count > 0 Then
                num += WOItems.Rows.Count
            End If
            If num > 0 Then
                num -= 1
            End If
            Dim LineItems(num) As cTaxCode.LineItem
            For i As Integer = 0 To Me.LineItems.Rows.Count - 1
                LineItems(i).Quantity = IIf(Me.LineItems.Rows(i).Item("quantity") Is DBNull.Value, 1, Me.LineItems.Rows(i).Item("quantity"))
                LineItems(i).Amount = IIf(Me.LineItems.Rows(i).Item("price") Is DBNull.Value, 0, Me.LineItems.Rows(i).Item("price"))
                LineItems(i).Taxable = IIf(Me.LineItems.Rows(i).Item("taxable") Is DBNull.Value, False, Me.LineItems.Rows(i).Item("taxable"))
            Next
            For i As Integer = 0 To WOItems.Rows.Count - 1
                LineItems(Me.LineItems.Rows.Count + i).Quantity = IIf(WOItems.Rows(i).Item("quantity") Is DBNull.Value, 1, WOItems.Rows(i).Item("quantity"))
                LineItems(Me.LineItems.Rows.Count + i).Amount = IIf(WOItems.Rows(i).Item("unit_price") Is DBNull.Value, 0, WOItems.Rows(i).Item("unit_price"))
                LineItems(Me.LineItems.Rows.Count + i).Taxable = IIf(WOItems.Rows(i).Item("taxable") Is DBNull.Value, False, WOItems.Rows(i).Item("taxable"))
            Next
            Return LineItems
        End Get
    End Property

    Public ReadOnly Property SalesTax() As Double
        Get
            If Me.TaxCode > 0 Then
                Dim TaxCode As New cTaxCode(Me.Database)
                TaxCode.Open(Me.TaxCode)
                Return TaxCode.CalculateTax(Me.LineItemsForSalesTaxCalculation, Me.Freight)
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property SalesTaxLineItems() As DataTable
        Get
            If Me.TaxCode > 0 Then
                Dim TaxCode As New cTaxCode(Me.Database)
                TaxCode.Open(Me.TaxCode)
                Return TaxCode.TaxPerAuthority(Me.LineItemsForSalesTaxCalculation, Me.Freight)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property WorkOrderItems() As DataTable
        Get
            Dim Sql As String = ""
            Sql &= "SELECT li.*, pn.prime, pn.item_type_id, ts.taxable, li.quantity * li.unit_price AS total_price"
            Sql &= " FROM work_order_item li"
            Sql &= " LEFT OUTER JOIN item_master pn ON pn.part_no=li.part_no"
            Sql &= " LEFT OUTER JOIN tax_status ts ON ts.id=li.tax_status_id"
            Sql &= " WHERE work_order_id IN"
            Sql &= " ( SELECT [id] FROM work_order WHERE service_order_id=@service_order_id)"
            Sql = Sql.Replace("@service_order_id", Me.Database.Escape(Me.OrderNo))
            Dim Table As DataTable = Me.Database.GetAll(Sql)
            If Me.Database.LastQuery.Successful Then
                Return Table
            Else
                Dim Err As String = Me.Database.LastQuery.ErrorMsg
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property Interactions() As DataTable
        Get
            Dim Sql As String = ""
            Sql &= " SELECT "
            Sql &= " contact_name, customer_no, id, entry_type_id, subject, memo, contact_id,"
            Sql &= " created_by, created_date, touch_date, touch_by, department, initiator,"
            Sql &= " date_last_updated, last_updated_by, ref_no"
            Sql &= " FROM journal"
            Sql &= " WHERE ref_no=" & Me.Database.Escape(Me.OrderNo)
            Sql &= " AND department=" & CInt(cInteraction.ReferenceTypes.Service)
            Return Me.Database.GetAll(Sql)
        End Get
    End Property

    Public ReadOnly Property RateTypes() As DataTable
        Get
            Return Me.Database.GetAll("SELECT id, name, hourly_equivalent FROM rate_type ORDER BY sort")
        End Get
    End Property

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
        ' Populate tables
        Me.PopulateOffices()
        Me.PopulateTaxCodes()
        Me.PopulateTruckTypes()
        Me.PopulateWorkType()
        Me.PopulateRateTables()
        Me.PopulateServiceOrderTypes()
        ' Blank Work orders table
        Me.WorkOrders = New DataTable
        Me.WorkOrders.Columns.Add("id")
        Me.WorkOrders.Columns.Add("equipment_id")
        Me.WorkOrders.Columns.Add("description")
        Me.WorkOrders.Columns.Add("problem_reported")
        Me.WorkOrders.Columns.Add("problem_found")
        Me.WorkOrders.Columns.Add("corrective_action")
        Me.WorkOrders.Columns("description").DefaultValue = ""
        Me.WorkOrders.Columns("problem_reported").DefaultValue = ""
        Me.WorkOrders.Columns("problem_found").DefaultValue = ""
        Me.WorkOrders.Columns("corrective_action").DefaultValue = ""
        ' Blank line items table
        Me.LineItems = New DataTable
        Me.LineItems.Columns.Add("id")
        Me.LineItems.Columns.Add("quantity")
        Me.LineItems.Columns.Add("part_no")
        Me.LineItems.Columns.Add("item_type_id")
        Me.LineItems.Columns.Add("serial_no")
        Me.LineItems.Columns.Add("equipment_id")
        Me.LineItems.Columns.Add("description")
        Me.LineItems.Columns.Add("station_id")
        Me.LineItems.Columns.Add("tax_status_id")
        Me.LineItems.Columns.Add("taxable")
        Me.LineItems.Columns.Add("prime")
        Me.LineItems.Columns.Add("price")
        Me.LineItems.Columns("part_no").DefaultValue = ""
        Me.LineItems.Columns("description").DefaultValue = ""
        Me.LineItems.Columns("serial_no").DefaultValue = ""
        Me.LineItems.Columns("item_type_id").DefaultValue = 0
        Me.LineItems.Columns("equipment_id").DefaultValue = 0
        Me.LineItems.Columns("station_id").DefaultValue = 0
        Me.LineItems.Columns("tax_status_id").DefaultValue = 2
        Me.LineItems.Columns("price").DefaultValue = 0
        Me.LineItems.Columns("quantity").DefaultValue = 0
        Me.LineItems.Columns("taxable").DefaultValue = False
        Me.LineItems.Columns("prime").DefaultValue = False
    End Sub

    Public Sub Open(ByVal Id As Integer)
        Dim SO As DataRow
        Dim Sql As String = ""
        Sql &= "SELECT so.*,"
        Sql &= " ship_to.cst_name AS company_name, ship_to.cst_addr1, ship_to.cst_addr2, ship_to.cst_city, ship_to.cst_state,"
        Sql &= " ship_to.cst_zip, ship_to.cst_phone, ship_to.cst_fax,"
        Sql &= " bill_to.cst_name AS bill_to_name, bill_to.cst_addr1 AS bill_to_address1, "
        Sql &= " bill_to.cst_addr2 AS bill_to_address2, bill_to.cst_city AS bill_to_city, bill_to.cst_state AS bill_to_state,"
        Sql &= " bill_to.cst_zip AS bill_to_zip, bill_to.cst_phone AS bill_to_phone, "
        Sql &= " bill_to.cst_fax AS bill_to_fax, bill_to.cst_stat AS terms,"
        Sql &= " (SELECT [name] FROM ADDRESS, pay_status WHERE cst_no=so.customer_no AND pay_status_id=pay_status.id) AS cst_stat"
        Sql &= " FROM service_order so"
        Sql &= " JOIN ADDRESS ship_to ON so.location_id=ship_to.cst_no"
        Sql &= " JOIN ADDRESS bill_to ON so.customer_no=bill_to.cst_no"
        Sql &= " WHERE so.id=@id"
        Sql = Sql.Replace("@id", Id)
        SO = Me.Database.GetRow(Sql)
        If Me.Database.LastQuery.RowsReturned = 1 Then
            Try
                Me._Id = Id
                Me.BillToNo = IIf(SO.Item("customer_no") Is DBNull.Value, 0, SO.Item("customer_no"))
                Me.ShipToNo = IIf(SO.Item("location_id") Is DBNull.Value, 0, SO.Item("location_id"))
                Me._ShipToName = SO.Item("company_name")
                Me.CalledInBy = Me.IsNull(SO.Item("caller_name"))
                Me.CalledInByID = Me.IsNull(SO.Item("caller_id"), 0)
                Me.ApprovedBy = SO.Item("approved_name")
                Me.ApprovedByID = Me.IsNull(SO.Item("approved_id"), 0)
                Me.ContactName = Me.IsNull(SO.Item("contact_name"))
                Me.ContactID = Me.IsNull(SO.Item("contact_id"), 0)
                Me.TakenBy = Me.IsNull(SO.Item("taken_by"))
                Me.DateDue = IIf(SO.Item("date_due") Is DBNull.Value, 0, SO.Item("date_due"))
                Me.DateCreated = IIf(SO.Item("date_created") Is DBNull.Value, 0, SO.Item("date_created"))
                Me.CalAgreement = IIf(SO.Item("cal_agreement_id") Is DBNull.Value, 0, SO.Item("cal_agreement_id"))
                Me.AgreementAdvancedNextDue = SO.Item("agreement_advanced_next")
                Me.Office = SO.Item("office")
                Me.Notes = Me.IsNull(SO.Item("notes"))
                Me.InternalNotes = Me.IsNull(SO.Item("additional_notes"))
                Me.EstimatedHours = SO.Item("estimated_hours")
                Me.WorkType = SO.Item("work_type_id")
                Me.ServiceOrderType = SO.Item("service_order_type")
                Me.Truck = IIf(SO.Item("truck_type_id") Is DBNull.Value, 0, SO.Item("truck_type_id"))
                Me.RateType = IIf(SO.Item("rate_type_id") Is DBNull.Value, 0, SO.Item("rate_type_id"))
                Me.RateTable = IIf(SO.Item("rate_table_id") Is DBNull.Value, 0, SO.Item("rate_table_id"))
                Me.Contract = IIf(SO.Item("contract") Is DBNull.Value, False, SO.Item("contract"))
                Me.Techs = SO.Item("techs") ' ?
                Me.Helpers = SO.Item("helpers") ' ?
                Me.AssignedTo = IIf(SO.Item("assigned_to") Is DBNull.Value, "", SO.Item("assigned_to"))
                Me.DateCompleted = IIf(SO.Item("date_completed") Is DBNull.Value, Nothing, SO.Item("date_completed"))
                Me.JobSummary = SO.Item("synopsis")
                Me.InvoiceNo = IIf(SO.Item("invoice_id") Is DBNull.Value, 0, SO.Item("invoice_id"))
                Me.PickUp = IIf(SO.Item("pick_up") Is DBNull.Value, False, SO.Item("pick_up"))
                Me.Po = Me.IsNull(SO.Item("po"))
                Me.Freight = SO.Item("freight")
                Me.TaxCode = IIf(SO.Item("tax_code_id") Is DBNull.Value, 0, SO.Item("tax_code_id"))
                Me.Shop = SO.Item("shop_work")
                Me.ChargeTo = SO.Item("charge_to")
                Me.ProjectId = IIf(SO.Item("project_id") Is DBNull.Value, 0, SO.Item("project_id"))
                Me.DownloadToPrimaryTechOnly = SO.Item("primary_tech_only")
                Me.StationId = SO.Item("station_id")
                Me.PriorityID = SO.Item("priority_id")
                Me.TechSignature = IIf(SO.Item("tech_signature") Is DBNull.Value, "", SO.Item("tech_signature"))
                Me.Voided = IIf(SO.Item("voided") Is DBNull.Value, False, SO.Item("voided"))
                Me.CreatedBy = Me.IsNull(SO.Item("created_by"))
                Me.DateCreated = SO.Item("date_created")
                Me.LastUpdateBy = Me.IsNull(SO.Item("last_updated_by"))
                Me.DateLastUpdated = IIf(SO.Item("date_last_updated") Is DBNull.Value, "", SO.Item("date_last_updated"))
            Catch ex As Exception
                Throw New Exception("Error opening service order: " & ex.ToString)
            End Try
            Try
                ' Ship To
                Dim co As New cCompany(Me.Database)
                co.Open(Me.ShipToNo)
                Me.ShipTo.Name = co.Name
                Me.ShipTo.Address1 = co.Address1
                Me.ShipTo.Address2 = co.Address2
                Me.ShipTo.City = co.City
                Me.ShipTo.State = co.State
                Me.ShipTo.Zip = co.Zip
                Me.ShipTo.Phone = co.Phone
                Me.ShipTo.Fax = co.Fax
                Me.ShipTo.Terms = co.Terms
                ' Bill to
                co = New cCompany(Me.Database)
                co.Open(Me.BillToNo)
                If co.BillingAddress1.Length > 0 Then
                    Me.BillTo.Name = co.BillingName
                    Me.BillTo.Address1 = co.BillingAddress1
                    Me.BillTo.Address2 = co.BillingAddress2
                    Me.BillTo.City = co.BillingCity
                    Me.BillTo.State = co.BillingState
                    Me.BillTo.Zip = co.BillingZip
                    Me.BillTo.Phone = co.BillingPhone
                    Me.BillTo.Fax = co.BillingFax
                    Me.BillTo.Terms = co.Terms
                Else
                    Me.BillTo.Name = co.Name
                    Me.BillTo.Address1 = co.Address1
                    Me.BillTo.Address2 = co.Address2
                    Me.BillTo.City = co.City
                    Me.BillTo.State = co.State
                    Me.BillTo.Zip = co.Zip
                    Me.BillTo.Phone = co.Phone
                    Me.BillTo.Fax = co.Fax
                    Me.BillTo.Terms = co.Terms
                End If
            Catch ex As Exception
                Throw New Exception("Error opening service order's company record: " & ex.ToString)
            End Try
            ' Get Work Orders
            Sql = "SELECT id, equipment_id, problem_reported, description, calibrated, problem_found, corrective_action"
            Sql &= " FROM work_order"
            Sql &= " WHERE service_order_id=@service_order_id"
            Sql = Sql.Replace("@service_order_id", Id)
            Me.WorkOrders = Me.Database.GetAll(Sql)
            ' Get Line Items
            Sql = "SELECT li.id, li.quantity, li.part_no, li.serial_no, li.equipment_id, li.description, li.price,"
            Sql &= " li.tax_status_id, li.station_id, pn.item_type_id, ts.taxable, pn.prime"
            Sql &= " FROM service_order_line_item li"
            Sql &= " LEFT OUTER JOIN item_master pn ON pn.part_no=li.part_no"
            Sql &= " LEFT OUTER JOIN tax_status ts ON ts.id=li.tax_status_id"
            Sql &= " WHERE service_order_id=" & Id
            Me.LineItems = Me.Database.GetAll(Sql)
            If Not Me.Database.LastQuery.Successful Then
                Dim er As String = Me.Database.LastQuery.ErrorMsg
                Throw New Exception(er)
            End If
            RaiseEvent Reload()
        Else
            If Not Me.Database.LastQuery.Successful Then
                Throw New Exception("Could not open service order: " & Me.Database.LastQuery.ErrorMsg)
            Else
                Throw New Exception("Service order #" & Id & " was not found.")
            End If
        End If
    End Sub

    Private Function IsNull(ByVal Value As Object, Optional ByVal ReturnValue As Object = "") As Object
        If Value Is DBNull.Value Then
            Return ReturnValue
        Else
            Return Value
        End If
    End Function

    Public Sub Save(Optional ByVal SaveWorkOrders As Boolean = True, Optional ByVal SaveLineItems As Boolean = True)
        Dim blnReturn As Boolean = False
        If Me._Id = Nothing Then
            Me.SaveNew()
            ' Create a new interaction
            Try
                If Me.CalledInBy.Trim.Length > 0 Then
                    Dim Sql As String = "INSERT INTO journal (contact_name, contact_id, customer_no, subject, memo, initiator, department, entry_type_id, created_by, created_date,"
                    Sql &= " touch_by, touch_date, date_last_updated, last_updated_by)"
                    Sql &= " VALUES (@contact_name, @contact_id, @customer_no, @subject, @notes, @initiator, @department, @type, @created_by,"
                    Sql &= " " & Me.Database.Timestamp & ", @employee, @when, " & Me.Database.Timestamp & ", @created_by, @agreement_advanced_next)"
                    ' Set parameters
                    Sql = Sql.Replace("@contact_name", Me.Database.Escape(Me.CalledInBy))
                    Sql = Sql.Replace("@contact_id", Me.Database.Escape(Me.CalledInByID))
                    Sql = Sql.Replace("@customer_no", Me.Database.Escape(Me.ShipToNo))
                    Sql = Sql.Replace("@employee", Me.Database.Escape(Me.TakenBy))
                    Sql = Sql.Replace("@when", Me.Database.Escape(Now))
                    Sql = Sql.Replace("@type", 1)
                    Sql = Sql.Replace("@subject", Me.Database.Escape("Service Order #" & Me.OrderNo))
                    Sql = Sql.Replace("@notes", Me.Database.Escape(Me.Notes))
                    Sql = Sql.Replace("@department", 2)
                    Sql = Sql.Replace("@initiator", 2)
                    Sql = Sql.Replace("@created_by", Me.Database.Escape(Me.LastUpdateBy))
                    Sql = Sql.Replace("@created_date", Me.Database.Escape(Now))
                    Me.Database.Execute(Sql)
                End If
            Catch
                ' Ignore
            End Try
        Else
            Me.SaveExisting()
        End If
        If SaveWorkOrders Then
            Me.SaveWorkOrders()
        End If
        If SaveLineItems Then
            Me.SaveLineItems()
        End If
        RaiseEvent Saved(Me)
        If SaveWorkOrders And SaveLineItems Then
            Me.Reopen()
        End If
    End Sub

    Public Sub Reopen()
        If Me.OrderNo > 0 Then
            Me.Open(Me.OrderNo)
        End If
    End Sub

    Private Sub SaveNew()
        Dim Sql As String = ""
        Sql &= "INSERT INTO service_order (contact_name, contact_id, location_id, office, cal_agreement_id, notes, date_due, "
        Sql &= " truck_type_id, rate_type_id, work_type_id, contract, techs, helpers, priority_id,"
        Sql &= " created_by, taken_by, date_created, pick_up, shop_work, po, caller_name, caller_id, "
        Sql &= " rate_table_id, date_last_updated, customer_no, tax_code_id, project_id, assigned_to, service_order_type,"
        Sql &= " estimated_hours, additional_notes, agreement_advanced_next, voided)"
        Sql &= " VALUES (@contact_name, @contact_id, @ship_to, @office, @cal_agreement_id, @notes, @date_due, @truck, @rate_type_id,"
        Sql &= " @work, @contract, @techs, @helpers, @priority,"
        Sql &= " @created_by, @taken_by, " & Me.Database.Timestamp & ", @pick_up, @shop, @po, @caller_name, @caller_id, @rate_table_id, "
        Sql &= " " & Me.Database.Timestamp & ", @bill_to, @tax_code_id, @project_id, @assigned_to, @service_order_type,"
        Sql &= " @estimated_hours, @internal_notes, @agreement_advanced_next, @voided)"
        Sql = Sql.Replace("@project_id", Me.ProjectId)
        Sql = Sql.Replace("@contact_name", Me.Database.Escape(Me.ContactName))
        Sql = Sql.Replace("@contact_id", Me.Database.Escape(Me.ContactID))
        Sql = Sql.Replace("@bill_to", Me.Database.Escape(Me.BillToNo))
        Sql = Sql.Replace("@ship_to", Me.Database.Escape(Me.ShipToNo))
        Sql = Sql.Replace("@office", Me.Database.Escape(Me.Office))
        Sql = Sql.Replace("@service_order_type", Me.ServiceOrderType)
        Sql = Sql.Replace("@cal_agreement_id", Me.CalAgreement)
        Sql = Sql.Replace("@notes", Me.Database.Escape(Me.Notes))
        Sql = Sql.Replace("@date_due", Me.Database.Escape(Me.DateDue))
        Sql = Sql.Replace("@truck", Me.Truck)
        Sql = Sql.Replace("@rate_type_id", Me.RateType)
        Sql = Sql.Replace("@work", Me.WorkType)
        Sql = Sql.Replace("@contract", Me.Database.Escape(Me.Contract))
        Sql = Sql.Replace("@techs", Me.Techs)
        Sql = Sql.Replace("@helpers", Me.Helpers)
        Sql = Sql.Replace("@priority", Me.PriorityID)
        Sql = Sql.Replace("@taken_by", Me.Database.Escape(Me.TakenBy))
        Sql = Sql.Replace("@estimated_hours", Me.EstimatedHours)
        Sql = Sql.Replace("@pick_up", Me.Database.Escape(Me.PickUp))
        Sql = Sql.Replace("@shop", Me.Database.Escape(Me.Shop))
        Sql = Sql.Replace("@po", Me.Database.Escape(Me.Po))
        Sql = Sql.Replace("@caller_name", Me.Database.Escape(Me.CalledInBy))
        Sql = Sql.Replace("@caller_id", Me.Database.Escape(Me.CalledInByID))
        Sql = Sql.Replace("@rate_table_id", Me.RateTable)
        Sql = Sql.Replace("@tax_code_id", Me.TaxCode)
        Sql = Sql.Replace("@assigned_to", Me.Database.Escape(Me.AssignedTo))
        Sql = Sql.Replace("@created_by", Me.Database.Escape(Me.LastUpdateBy))
        Sql = Sql.Replace("@internal_notes", Me.Database.Escape(Me.InternalNotes))
        Sql = Sql.Replace("@agreement_advanced_next", Me.Database.Escape(Me.AgreementAdvancedNextDue))
        Sql = Sql.Replace("@voided", Me.Database.Escape(Me.Voided))
        Me.Database.InsertAndReturnId(Sql)
        If Me.Database.LastQuery.Successful Then
            Me._Id = Me.Database.LastQuery.InsertId
        Else
            Throw New Exception(Me.Database.LastQuery.ErrorMsg)
        End If
    End Sub

    Private Sub SaveExisting()
        Dim Sql As String = ""
        Sql &= "UPDATE service_order SET"
        Sql &= " contact_name=@contact_name,"
        Sql &= " contact_id=@contact_id,"
        Sql &= " customer_no=@bill_to,"
        Sql &= " location_id=@ship_to,"
        Sql &= " office=@office,"
        Sql &= " cal_agreement_id=@cal_agreement_id,"
        Sql &= " notes=@notes,"
        Sql &= " date_due=@date_due,"
        Sql &= " date_completed=@date_completed,"
        Sql &= " invoice_id=@invoice_id,"
        Sql &= " truck_type_id=@truck,"
        Sql &= " work_type_id=@work,"
        Sql &= " contract=@contract,"
        Sql &= " approved_name=@approved_name,"
        Sql &= " approved_id=@approved_id,"
        Sql &= " synopsis=@job_summary,"
        Sql &= " pick_up=@pick_up,"
        Sql &= " po=@po,"
        Sql &= " caller_name=@caller_name,"
        Sql &= " caller_id=@caller_id,"
        Sql &= " additional_notes=@internal_notes,"
        Sql &= " rate_table_id=@rate_table_id,"
        Sql &= " freight=" & Me.Database.ToCurrency("@freight") & ","
        Sql &= " date_last_updated=" & Me.Database.Timestamp & ","
        Sql &= " project_id=@project_id,"
        Sql &= " taken_by=@taken_by,"
        Sql &= " tax_code_id=@tax_code_id,"
        Sql &= " assigned_to=@assigned_to,"
        Sql &= " service_order_type=@service_order_type,"
        Sql &= " charge_to=@charge_to,"
        Sql &= " estimated_hours=@estimated_hours,"
        Sql &= " primary_tech_only=@primary_tech_only,"
        Sql &= " station_id=@station_id,"
        Sql &= " shop_work=@shop,"
        Sql &= " priority_id=@priority,"
        Sql &= " agreement_advanced_next=@agreement_advanced_next,"
        Sql &= " voided=@voided"
        Sql &= " WHERE id=@id"
        Sql = Sql.Replace("@id", Me.OrderNo)
        Sql = Sql.Replace("@project_id", Me.ProjectId)
        Sql = Sql.Replace("@bill_to", Me.Database.Escape(Me.BillToNo))
        Sql = Sql.Replace("@ship_to", Me.Database.Escape(Me.ShipToNo))
        Sql = Sql.Replace("@caller_name", Me.Database.Escape(Me.CalledInBy))
        Sql = Sql.Replace("@contact_name", Me.Database.Escape(Me.ContactName))
        Sql = Sql.Replace("@contact_id", Me.Database.Escape(Me.ContactID))
        Sql = Sql.Replace("@caller_id", Me.Database.Escape(Me.CalledInByID))
        Sql = Sql.Replace("@office", Me.Database.Escape(Me.Office))
        Sql = Sql.Replace("@service_order_type", Me.ServiceOrderType)
        Sql = Sql.Replace("@shop", Me.Database.Escape(Me.Shop))
        Sql = Sql.Replace("@taken_by", Me.Database.Escape(Me.TakenBy))
        Sql = Sql.Replace("@po", Me.Database.Escape(Me.Po))
        Sql = Sql.Replace("@notes", Me.Database.Escape(Me.Notes))
        Sql = Sql.Replace("@truck", Me.Truck)
        Sql = Sql.Replace("@rate_type", Me.RateType)
        Sql = Sql.Replace("@work", Me.WorkType)
        Sql = Sql.Replace("@techs", Me.Techs)
        Sql = Sql.Replace("@helpers", Me.Helpers)
        Sql = Sql.Replace("@priority", Me.PriorityID)
        Sql = Sql.Replace("@contract", Me.Database.Escape(Me.Contract))
        Sql = Sql.Replace("@pick_up", Me.Database.Escape(Me.PickUp))
        Sql = Sql.Replace("@rate_table_id", Me.RateTable)
        Sql = Sql.Replace("@date_due", Me.Database.Escape(Me.DateDue))
        Sql = Sql.Replace("@cal_agreement_id", Me.CalAgreement)
        Sql = Sql.Replace("@estimated_hours", Me.EstimatedHours)
        Sql = Sql.Replace("@tax_code_id", Me.TaxCode)
        Sql = Sql.Replace("@assigned_to", Me.Database.Escape(Me.AssignedTo))
        Sql = Sql.Replace("@internal_notes", Me.Database.Escape(Me.InternalNotes))
        Sql = Sql.Replace("@date_completed", Me.Database.Escape(IIf(Me.DateCompleted = Nothing, DBNull.Value, Me.DateCompleted)))
        Sql = Sql.Replace("@approved_name", Me.Database.Escape(Me.ApprovedBy))
        Sql = Sql.Replace("@approved_id", Me.Database.Escape(Me.ApprovedByID))
        Sql = Sql.Replace("@invoice_id", Me.InvoiceNo)
        Sql = Sql.Replace("@freight", Me.Freight)
        Sql = Sql.Replace("@job_summary", Me.Database.Escape(Me.JobSummary))
        Sql = Sql.Replace("@charge_to", Me.ChargeTo)
        Sql = Sql.Replace("@primary_tech_only", Me.Database.Escape(Me.DownloadToPrimaryTechOnly))
        Sql = Sql.Replace("@station_id", Me.StationId)
        Sql = Sql.Replace("@agreement_advanced_next", Me.AgreementAdvancedNextDue)
        Sql = Sql.Replace("@voided", Me.Database.Escape(Me.Voided))
        Me.Database.Execute(Sql)
        RaiseEvent Progress(50, "Saving Work Orders")
        If Me.Database.LastQuery.Successful Then
            ' Save work orders
        Else
            Throw New Exception(Me.Database.LastQuery.ErrorMsg)
        End If
    End Sub

    Public Sub SaveWorkOrders()
        Dim Count As Integer = 0
        Dim Sql As String = ""
        For Each Row As DataRow In Me.WorkOrders.Rows
            Select Case Row.RowState
                Case DataRowState.Modified
                    Sql = "UPDATE work_order"
                    Sql &= " SET equipment_id=@equipment_id,"
                    Sql &= " description=@description, problem_reported=@problem_reported,"
                    Sql &= " date_last_updated = " & Me.Database.Timestamp & ""
                    Sql &= " WHERE id=@work_order_id"
                    Sql = Sql.Replace("@work_order_id", Row.Item("id"))
                    Sql = Sql.Replace("@description", Me.Database.Escape(Row.Item("description")))
                    Sql = Sql.Replace("@problem_reported", Me.Database.Escape(Row.Item("problem_reported")))
                    Sql = Sql.Replace("@equipment_id", Row.Item("equipment_id"))
                    Me.Database.Execute(Sql)
                    If Not Me.Database.LastQuery.Successful Then
                        Dim Err As String = Me.Database.LastQuery.ErrorMsg
                    End If
                Case DataRowState.Added
                    'If Row.Item("description") IsNot DBNull.Value And Row.Item("problem_reported") IsNot DBNull.Value Then
                    'If Row.Item("description") <> "" Or Row.Item("problem_reported") <> "" Then
                    Sql = "INSERT INTO work_order (service_order_id, equipment_id, description, problem_reported, date_last_updated)"
                    Sql &= " VALUES (@service_order_id, @equipment_id, @description, @problem_reported, " & Me.Database.Timestamp & ")"
                    Sql = Sql.Replace("@service_order_id", Me._Id)
                    Sql = Sql.Replace("@equipment_id", IIf(Row.Item("equipment_id") Is DBNull.Value, 0, Row.Item("equipment_id")))
                    Sql = Sql.Replace("@description", Me.Database.Escape(IIf(Row.Item("description") Is DBNull.Value, "", Row.Item("description"))))
                    Sql = Sql.Replace("@problem_reported", Me.Database.Escape(IIf(Row.Item("problem_reported") Is DBNull.Value, "", Row.Item("problem_reported"))))
                    Me.Database.InsertAndReturnId(Sql)
                    If Not Me.Database.LastQuery.Successful Then
                        Dim Err As String = Me.Database.LastQuery.ErrorMsg
                    End If
                    'End If
                    'End If
                Case DataRowState.Deleted
                    ' *** Does this work????
                    If Row.Item("id") IsNot DBNull.Value Then
                        Me.Database.Execute("DELETE FROM work_order WHERE id=" & Row.Item("id"))
                    End If
            End Select
            Count += 1
        Next
    End Sub

    Public Sub SaveLineItems()
        For Each Row As DataRow In Me.LineItems.Rows
            ' Catch weird null error
            If Row.Item("quantity") Is DBNull.Value Then
                Row.Item("quantity") = 0
            ElseIf Row.Item("quantity") = Nothing Then
                Row.Item("quantity") = 0
            End If
            ' Continue
            Select Case Row.RowState
                Case DataRowState.Modified
                    If Row.Item("quantity") > 0 Then
                        Dim Sql As String = ""
                        Sql = "UPDATE service_order_line_item SET"
                        Sql &= " quantity=" & Me.Database.Escape(Row.Item("quantity"))
                        Sql &= ", part_no=" & Me.Database.Escape(Row.Item("part_no"))
                        Sql &= ", serial_no=" & Me.Database.Escape(Row.Item("serial_no"))
                        Sql &= ", equipment_id=" & Me.Database.Escape(Row.Item("equipment_id"))
                        Sql &= ", description=" & Me.Database.Escape(Row.Item("description"))
                        Sql &= ", price=" & Me.Database.Escape(Row.Item("price"))
                        Sql &= ", tax_status_id=" & Me.Database.Escape(Row.Item("tax_status_id"))
                        Sql &= ", station_id=" & Me.Database.Escape(Row.Item("station_id"))
                        Sql &= ", date_last_updated=" & Me.Database.Escape(Now)
                        Sql &= " WHERE id=" & Row.Item("id")
                        Me.Database.Execute(Sql)
                    Else
                        Dim Sql As String = ""
                        Sql = "DELETE FROM service_order_line_item WHERE id=" & Row.Item("id")
                        Me.Database.Execute(Sql)
                    End If
                Case DataRowState.Added
                    If Row.Item("quantity") > 0 Then
                        Dim Sql As String = ""
                        Sql = "INSERT INTO service_order_line_item"
                        Sql &= " (service_order_id, quantity, part_no, serial_no, equipment_id, description,"
                        Sql &= " price, tax_status_id, station_id, date_last_updated)"
                        Sql &= " VALUES (" & Me.OrderNo
                        Sql &= ", " & Me.Database.Escape(Me.IsNull(Row.Item("quantity"), 1))
                        Sql &= ", " & Me.Database.Escape(Row.Item("part_no"))
                        Sql &= ", " & Me.Database.Escape(Me.IsNull(Row.Item("serial_no")))
                        Sql &= ", " & Me.Database.Escape(Me.IsNull(Row.Item("equipment_id"), 0))
                        Sql &= ", " & Me.Database.Escape(Row.Item("description"))
                        Sql &= ", " & Me.Database.ToCurrency(Me.Database.Escape(Row.Item("price")))
                        Sql &= ", " & Me.Database.Escape(Me.IsNull(Row.Item("tax_status_id"), 1))
                        Sql &= ", " & Me.Database.Escape(IIf(Row.Item("station_id") Is DBNull.Value, 0, Row.Item("station_id")))
                        Sql &= ", " & Me.Database.Escape(Now)
                        Sql &= ")"
                        Me.Database.InsertAndReturnId(Sql)
                        If Not Me.Database.LastQuery.Successful Then
                            Throw New Exception(Me.Database.LastQuery.ErrorMsg)
                        End If
                    End If
                Case DataRowState.Deleted
                    If Row.Item("id") IsNot DBNull.Value Then
                        Me.Database.Execute("DELETE FROM service_order_line_item WHERE id=" & Row.Item("id"))
                    End If
            End Select
        Next
    End Sub

    Public Function Employees(Optional ByVal All As Boolean = True) As DataTable
        If All Then
            Return Me.Database.GetAll("SELECT windows_user, (last_name + ', ' + first_name) AS name, deactivated FROM employee ORDER BY last_name, first_name")
        Else
            Return Me.Database.GetAll("SELECT windows_user, (last_name + ', ' + first_name) AS name, deactivated FROM employee WHERE deactivated=0 ORDER BY last_name, first_name")
        End If
    End Function

    Private Sub PopulateTruckTypes()
        Me.TruckTypes = Me.Database.GetAll("SELECT id, name FROM truck_type ORDER BY sort")
    End Sub

    Private Sub PopulateRateTables()
        Me.RateTables = Me.Database.GetAll("SELECT id, name, method FROM rate_table")
    End Sub

    Private Sub PopulateWorkType()
        Me.WorkTypes = Me.Database.GetAll("SELECT id, name FROM work_type ORDER BY sort")
    End Sub

    Private Sub PopulateOffices()
        Me.Offices = Me.Database.GetAll("SELECT number, name FROM office ORDER BY sort")
    End Sub

    Private Sub PopulateTaxCodes()
        Me.TaxCodes = Me.Database.GetAll("SELECT id, name FROM tax_code ORDER BY name")
    End Sub

    Private Sub PopulateServiceOrderTypes()
        Me.ServiceOrderTypes = Me.Database.GetAll("SELECT id, name FROM service_order_type ORDER BY sort")
    End Sub

    Public Function TechList() As String()
        Dim Table As DataTable = Me.Database.GetAll("SELECT DISTINCT owner FROM schedule WHERE reference_id=" & Me.OrderNo & " AND owner <> " & Me.Database.Escape(Me.AssignedTo))
        Dim Out(Table.Rows.Count + 1) As String
        Out(0) = Me.AssignedTo
        For i As Integer = 1 To Table.Rows.Count
            Out(i) = Table.Rows(i - 1).Item("owner").ToString
        Next
        Return Out
    End Function

    Public Function GetWorkOrderRow(ByVal Id As Integer) As DataRow
        For Each r As DataRow In Me.WorkOrders.Rows
            If r.Item("id") IsNot DBNull.Value Then
                If r.Item("id") = Id Then
                    Return r
                End If
            End If
        Next
        Return Nothing
    End Function

    Public Function GetPurchaseOrders() As DataTable
        Dim Sql As String = ""
        Sql &= "SELECT po.po_no, po.vendor_no, c.cst_name AS vendor_name, "
        Sql &= " po.po_date, po.date_ordered, po.date_planned_ship, po.date_expected,"
        Sql &= " ((SELECT SUM(quantity*unit_price) FROM purchase_order_item WHERE po_no=po.po_no) + shipping_charge + tax) AS total_price,"
        Sql &= " po.office, po.requested_by,"
        Sql &= " type = CASE po.rma WHEN 1 THEN 'RMA' ELSE 'PO' END"
        Sql &= " FROM purchase_order po"
        Sql &= " LEFT JOIN ADDRESS c ON po.vendor_no=c.cst_no"
        Sql &= " WHERE po.type=3 AND po.our_order_no=" & Me.Database.Escape(Me.OrderNo)
        Sql &= " ORDER BY po_date"
        Dim Table As DataTable = Me.Database.GetAll(Sql)
        If Not Me.Database.LastQuery.Successful Then
            MsgBox(Me.Database.LastQuery.ErrorMsg & " " & Sql)
        End If
        Return Table
    End Function

    Public Function SurchargeTypes() As DataTable
        Return Me.Database.GetAll("SELECT id, name, price FROM surcharge WHERE deleted=0")
    End Function

    Public Function TimesheetEntries() As DataTable
        Dim Sql As String = ""
        Sql = "SELECT t.*, c.name AS category_name, c.color AS category_color"
        Sql &= " FROM timesheet t"
        Sql &= " LEFT JOIN schedule_category c ON t.category_id=c.id"
        Sql &= " WHERE t.reference_id = " & Me._Id & " AND t.deleted=0"
        Sql &= " ORDER BY date_start ASC, owner ASC"
        Return Me.Database.GetAll(Sql)
    End Function

    Public Function GetTechniciansFromTimesheet() As DataTable
        Dim Sql As String = ""
        Sql = "SELECT DISTINCT t.owner AS tech "
        Sql &= " FROM timesheet t"
        Sql &= " WHERE t.reference_id = " & Me._Id & " AND t.deleted=0"
        Return Me.Database.GetAll(Sql)
    End Function

    Public Function ToGravityDocument(ByVal Template As String) As GravityDocument.gDocument
        ' If no template specified
        If Template.Length = 0 Then
            Dim id As Integer = Me.Database.GetOne("SELECT value FROM settings WHERE property='Template SO Sig'")
            Template = Me.Database.GetOne("SELECT html FROM template WHERE id=" & id)
        End If
        ' Create Gravity Document
        Dim Doc As New GravityDocument.gDocument(Me.Database.GetOne("SELECT value FROM settings WHERE property='Page Height in Pixels'"))
        Doc.LoadXml(Template)
        ' Settings
        Doc.FormType = GravityDocument.gDocument.FormTypes.ServiceOrder
        Doc.ReferenceID = Me.OrderNo
        Dim Page As GravityDocument.gPage = Doc.GetPage(1)
        ' Ship To
        Page.AddVariable("%customer_no%", Me.ShipToNo)
        Page.AddVariable("%company_name%", Me.ShipTo.Name)
        If Not Me.ShipTo.Address2.Length = 0 Then
            Page.AddVariable("%address%", Me.ShipTo.Address1 & ControlChars.CrLf & Me.ShipTo.Address2)
        Else
            Page.AddVariable("%address%", Me.ShipTo.Address1)
        End If
        Page.AddVariable("%city%", Me.ShipTo.City)
        Page.AddVariable("%state%", Me.ShipTo.State)
        Page.AddVariable("%zip%", Me.ShipTo.Zip)
        Page.AddVariable("%phone%", Me.ShipTo.Phone)
        Page.AddVariable("%fax%", Me.ShipTo.Fax)
        Page.AddVariable("%ship_to_no%", Me.ShipToNo)
        ' Bill To
        Dim BillToCo As New cCompany(Me.Database)
        BillToCo.Open(Me.BillToNo)
        Page.AddVariable("%bill_to_no%", Me.BillToNo)
        Page.AddVariable("%bill_to_name%", BillToCo.Name)
        ' If bill to has a billing address use it
        Dim BillToAddress As String = ""
        If BillToCo.BillingAddress1.Length > 0 Then
            BillToAddress = BillToCo.BillingAddress1
            If BillToCo.BillingAddress2.Length > 0 Then
                BillToAddress &= ControlChars.CrLf & BillToCo.BillingAddress2
            End If
        Else
            BillToAddress = BillToCo.Address1
            If BillToCo.Address2.Length > 0 Then
                BillToAddress &= ControlChars.CrLf & BillToCo.Address2
            End If
        End If
        If BillToCo.BillingName.Length > 0 Then
            Page.AddVariable("%bill_to_name%", BillToCo.BillingName)
            Page.AddVariable("%bill_to_address%", BillToAddress)
            Page.AddVariable("%bill_to_city%", BillToCo.BillingCity)
            Page.AddVariable("%bill_to_state%", BillToCo.BillingState)
            Page.AddVariable("%bill_to_zip%", BillToCo.BillingZip)
        Else
            Page.AddVariable("%bill_to_name%", BillToCo.Name)
            Page.AddVariable("%bill_to_address%", BillToAddress)
            Page.AddVariable("%bill_to_city%", BillToCo.City)
            Page.AddVariable("%bill_to_state%", BillToCo.State)
            Page.AddVariable("%bill_to_zip%", BillToCo.Zip)
        End If
        Page.AddVariable("%bill_to_phone%", BillToCo.Phone)
        Page.AddVariable("%bill_to_fax%", BillToCo.Fax)
        Page.AddVariable("%tax_exempt_no%", BillToCo.TaxNo)
        If BillToCo.TaxExemptThrough <> Nothing Then
            Page.AddVariable("%tax_exempt_thru%", BillToCo.TaxExemptThrough)
            Page.AddVariable("%tax_exempt_yn%", IIf(BillToCo.TaxExemptThrough > Now, "Y", "N"))
        Else
            Page.AddVariable("%tax_exempt_thru%", "--")
            Page.AddVariable("%tax_exempt_yn%", "N")
        End If
        ' Our office
        If Me.Office > 0 Then
            Dim Company As New cCompany(Me.Database)
            Try
                Dim OurAddress As String = Company.Address1
                If Company.Address2.Length > 0 Then
                    OurAddress &= ControlChars.CrLf & Company.Address2
                End If
                Company.Open(Me.Office)
                Page.AddVariable("%office_name%", Company.Name)
                Page.AddVariable("%office_name%", OurAddress)
                Page.AddVariable("%office_city%", Company.City)
                Page.AddVariable("%office_state%", Company.State)
                Page.AddVariable("%office_zip%", Company.Zip)
                Page.AddVariable("%office_phone%", Company.Phone)
                Page.AddVariable("%office_fax%", Company.Fax)
                Page.AddVariable("%office_website%", Company.WebSite)
                Page.AddVariable("%office_country%", Company.Country)
                Page.AddVariable("%office_email%", Company.APEmailAddress)
            Catch ex As Exception
                Page.AddVariable("%office_name%", "")
                Page.AddVariable("%office_name%", "")
                Page.AddVariable("%office_city%", "")
                Page.AddVariable("%office_state%", "")
                Page.AddVariable("%office_zip%", "")
                Page.AddVariable("%office_phone%", "")
                Page.AddVariable("%office_fax%", "")
                Page.AddVariable("%office_website%", "")
                Page.AddVariable("%office_country%", "")
                Page.AddVariable("%office_email%", "")
            End Try
        Else
            Page.AddVariable("%office_name%", "")
            Page.AddVariable("%office_name%", "")
            Page.AddVariable("%office_city%", "")
            Page.AddVariable("%office_state%", "")
            Page.AddVariable("%office_zip%", "")
            Page.AddVariable("%office_phone%", "")
            Page.AddVariable("%office_fax%", "")
            Page.AddVariable("%office_website%", "")
            Page.AddVariable("%office_country%", "")
            Page.AddVariable("%office_email%", "")
        End If
        ' Contact
        If Me.ContactID > 0 Then
            Try
                Dim Contact As New cContact(Me.Database)
                Contact.Open(Me.ContactID)
                Dim ContactAddress As String = Contact.Address1
                If Contact.Address2.Length > 0 Then
                    ContactAddress &= ControlChars.CrLf & Contact.Address2
                End If
                Page.AddVariable("%contact_fname%", Contact.FirstName)
                Page.AddVariable("%contact_lname%", Contact.LastName)
                Page.AddVariable("%contact_title%", Contact.Title)
                Page.AddVariable("%contact_phone%", Contact.BusinessPhone)
                Page.AddVariable("%contact_cell%", Contact.CellPhone)
                Page.AddVariable("%contact_city%", Contact.City)
                Page.AddVariable("%contact_state%", Contact.State)
                Page.AddVariable("%contact_zip%", Contact.ZipCode)
                Page.AddVariable("%contact_title%", Contact.Title)
                Page.AddVariable("%contact_dept%", Contact.Department)
                Page.AddVariable("%contact_salutation%", Contact.Salutation)
                Page.AddVariable("%contact_fax%", Contact.Fax)
                Page.AddVariable("%contact_email%", Contact.Email)
                Page.AddVariable("%contact_address%", ContactAddress)
            Catch ex As Exception
                Page.AddVariable("%contact_fname%", "")
                Page.AddVariable("%contact_lname%", "")
                Page.AddVariable("%contact_title%", "")
                Page.AddVariable("%contact_phone%", "")
                Page.AddVariable("%contact_cell%", "")
                Page.AddVariable("%contact_city%", "")
                Page.AddVariable("%contact_state%", "")
                Page.AddVariable("%contact_zip%", "")
                Page.AddVariable("%contact_title%", "")
                Page.AddVariable("%contact_dept%", "")
                Page.AddVariable("%contact_salutation%", "")
                Page.AddVariable("%contact_fax%", "")
                Page.AddVariable("%contact_email%", "")
                Page.AddVariable("%contact_address%", "")
            End Try
        Else
            Page.AddVariable("%contact_fname%", "")
            Page.AddVariable("%contact_lname%", "")
            Page.AddVariable("%contact_title%", "")
            Page.AddVariable("%contact_phone%", "")
            Page.AddVariable("%contact_cell%", "")
            Page.AddVariable("%contact_city%", "")
            Page.AddVariable("%contact_state%", "")
            Page.AddVariable("%contact_zip%", "")
            Page.AddVariable("%contact_title%", "")
            Page.AddVariable("%contact_dept%", "")
            Page.AddVariable("%contact_salutation%", "")
            Page.AddVariable("%contact_fax%", "")
            Page.AddVariable("%contact_email%", "")
            Page.AddVariable("%contact_address%", "")
        End If
        ' Terms
        Dim Terms As String = ""
        If Me.BillTo.Terms IsNot DBNull.Value Then
            Terms = Me.Database.GetOne("SELECT name FROM pay_status WHERE id=" & Me.BillTo.Terms)
        End If
        Page.AddVariable("%terms%", Terms)

        ' Service Order Details
        Page.AddVariable("%service_order_id%", Me.OrderNo)
        Page.AddVariable("%caller%", Me.CalledInBy)
        Page.AddVariable("%received_by%", Me.TakenBy)
        Page.AddVariable("%contact%", Me.ContactName)
        Page.AddVariable("%approved_by%", Me.ApprovedBy)
        Page.AddVariable("%date_received%", Format(Me.DateCreated, "MM/dd/yyyy"))
        Page.AddVariable("%date_due%", Format(Me.DateDue, "MM/dd/yyyy"))
        Page.AddVariable("%po%", Me.Po)
        If Me.DateScheduled = Nothing Then
            Page.AddVariable("%date_scheduled%", "--")
        Else
            Page.AddVariable("%date_scheduled%", Format(Me.DateScheduled, "MM/dd/yyyy"))
        End If
        If Me.DateCompleted = Nothing Then
            Page.AddVariable("%date_completed%", "")
            Page.AddVariable("%date%", Now.ToString("MM/dd/yyyy"))
        Else
            Page.AddVariable("%date_completed%", Format(Me.DateCompleted, "MM/dd/yyyy"))
            Page.AddVariable("%date%", Format(Me.DateCompleted, "MM/dd/yyyy"))
        End If
        Page.AddVariable("%notes%", Me.Notes)
        Page.AddVariable("%zone%", "")
        Page.AddVariable("%hourly_rate%", "")
        Page.AddVariable("%job_summary%", Me.JobSummary)
        ' Approved
        Page.AddVariable("%approved_name%", Me.ApprovedBy)
        If Not Me.DateCompleted = Nothing Then
            Page.AddVariable("%date%", Me.DateCompleted)
        Else
            If Me.DateCompleted = Nothing Then
                Page.AddVariable("%date%", Format(Now, "MM/dd/yyyy"))
            Else
                Page.AddVariable("%date%", Format(Me.DateCompleted, "MM/dd/yyyy"))
            End If
        End If
        ' Next Due Date
        If Me.CalAgreement = 0 Then
            Page.AddVariable("%next_cal_date%", "N/A")
        Else
            Dim Agreement As New cCalAgreement(Me.Database)
            Agreement.Open(Me.CalAgreement)
            Page.AddVariable("%next_cal_date%", Agreement.DateNextCal)
        End If

        ' Techs
        Dim Users As String = Me.Database.Escape(Me.AssignedTo) & ", "
        Dim strTechs As String = ""
        Dim r As DataRow
        Dim dt As DataTable = Me.Database.GetAll("SELECT DISTINCT owner FROM schedule WHERE reference_id=" & Me.OrderNo)
        For Each r In dt.Rows
            If Not r.Item("owner") = Me.AssignedTo Then
                Users &= Me.Database.Escape(r.Item("owner")) & ", "
            End If
        Next
        Users = Users.Substring(0, Users.Length - 2)
        Dim Techs As DataTable = Me.Database.GetAll("SELECT (last_name + ', ' + first_name) AS display_name FROM employee WHERE windows_user IN (" & Users & ")")
        For Each r In Techs.Rows
            strTechs &= r.Item("display_name") & "; "
        Next
        If strTechs.Length > 2 Then
            strTechs = strTechs.Substring(0, strTechs.Length - 2)
        End If
        Page.AddVariable("%techs%", strTechs)
        Page.AddVariable("%lead_tech%", Me.AssignedTo)

        Page.AddVariable("%hours_on_job%", Me.Database.GetOne("SELECT SUM(duration) FROM schedule WHERE category_id=1 AND reference_id=" & Me.OrderNo))

        Dim WorkOrderTotal As Double = Me.WorkOrderItemsTotal
        Page.AddVariable("%wo_subtotal%", Format(WorkOrderTotal, "$#.00"))
        Page.AddVariable("%wo_parts", Format(Me.WorkOrderPartsTotal, "$#.00"))
        Page.AddVariable("%wo_additional", Format(Me.WorkOrderAdditionalTotal, "$#.00"))

        Page.AddVariable("%so_subtotal%", Format(Me.LineItemsTotal, "$#.00"))
        Page.AddVariable("%so_parts%", Format(Me.PartsTotal, "$#.00"))
        Page.AddVariable("%trip_fee%", Format(Me.TravelTotal, "$#.00"))
        Page.AddVariable("%labor_charge%", Format(Me.LaborTotal, "$#.00"))
        Page.AddVariable("%trip_labor%", Format(Me.LaborTotal + Me.TravelTotal, "$#.00"))
        Page.AddVariable("%surcharges%", Format(Me.SurchargesTotal, "$#.00"))
        Page.AddVariable("%discounts%", Format(Me.DiscountsTotal, "$#.00"))
        Page.AddVariable("%special%", Format(Me.SpecialTotal, "$#.00"))
        Page.AddVariable("%freight%", Format(Me.Freight, "$#.00"))
        Page.AddVariable("%sur_plus_spec%", Format(Me.SurchargesTotal + Me.SpecialTotal, "$#.00"))

        Page.AddVariable("%parts%", Format(Me.PartsTotal + Me.WorkOrderPartsTotal, "$#.00"))
        Page.AddVariable("%misc%", Format(Me.WorkOrderAdditionalTotal + Me.SurchargesTotal + Me.SpecialTotal, "$#.00"))

        Page.AddVariable("%additional_total%", Format(IIf(WorkOrderTotal = Nothing, 0, WorkOrderTotal), "$#.00"))
        Page.AddVariable("%wo_parts_sur%", Format(Me.WorkOrderItemsTotal + Me.PartsTotal + Me.SurchargesTotal, "$#.00"))
        Page.AddVariable("%wo_plus_parts%", Format(Me.WorkOrderItemsTotal + Me.PartsTotal, "$#.00"))

        Dim SubTotal As Double = Me.LineItemsTotal + Me.WorkOrderItemsTotal - Me.OtherTaxTotal
        Page.AddVariable("%subtotal%", Format(SubTotal + Me.Freight, "$#.00"))
        Page.AddVariable("%subtotal_no_freight%", Format(SubTotal, "$#.00"))
        Page.AddVariable("%other_taxes%", Format(Me.OtherTaxTotal, "$#.00"))
        Page.AddVariable("%tax%", Format(Me.SalesTax, "$#.00"))
        Page.AddVariable("%tax_subtotal%", Format(Me.SalesTax + Me.OtherTaxTotal, "$#.00"))
        Page.AddVariable("%total%", Format(SubTotal + Me.Freight + Me.SalesTax, "$#.00"))


        Dim WOTable As GravityDocument.gElement = Page.GetTableBySource("work_orders")
        If WOTable IsNot Nothing Then
            WOTable.Table.Data = Me.WorkOrders
        End If
        Return Doc
    End Function

End Class
