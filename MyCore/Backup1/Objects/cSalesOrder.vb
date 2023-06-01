Imports MyCore.Data

Public Class cSalesOrder

    Dim _Id As Integer = 0
    Dim _TaxCodeId As Integer = 0

    Public BillTo As String = ""
    Public ShipTo As String = ""
    Public ProjectId As Integer = 0
    Public Salesperson As String = ""
    Public CustomerPO As String = ""
    Public DateCreated As Date = Now
    Public DateLastUpdated As Date = Now
    Public DateDue As Date = Now
    Public DateDelivered As Date = Nothing
    Public ShipVia As Integer = 0
    Public TrackingNo As String = ""
    Public QuoteId As Integer = 0
    Public LeadId As Integer = 0
    Public Contact As String = ""
    Public ContactID As Integer = 0
    Public Office As Integer = 0
    Public Notes As String = ""
    Public CreatedBy As String = ""
    Public LastUpdatedBy As String = ""
    Public Technician As String = ""
    Public Fob As String = ""
    Public Freight As Double = 0
    Public InvoiceNo As String = ""
    Public TermsId As Integer = 0
    Public ServiceOrderId As Integer = 0
    Public InternalNotes As String = ""
    Public NotifyStateStatus As Integer = 0
    Public Voided As Boolean = False

    Public LineItems As DataTable
    Public Offices As DataTable
    Public ShipVias As DataTable
    Public Employees As DataTable
    Public TaxCodes As DataTable

    Public TaxCode As cTaxCode

    Public Database As MyCore.Data.EasySql
    Public Event Reload()
    Public Event Saved(ByVal SalesOrder As cSalesOrder)

    Public Property TaxCodeId() As Integer
        Get
            Return Me._TaxCodeId
        End Get
        Set(ByVal value As Integer)
            If value > 0 Then
                Me.TaxCode.Open(value)
            End If
            Me._TaxCodeId = value
        End Set
    End Property

    Public ReadOnly Property OrderNo() As Integer
        Get
            Return Me._Id
        End Get
    End Property

    Public ReadOnly Property SalesTax() As Double
        Get
            If Me.TaxCodeId > 0 Then
                ' Line Items
                Dim LineItems(Me.LineItems.Rows.Count - 1) As cTaxCode.LineItem
                For i As Integer = 0 To Me.LineItems.Rows.Count - 1
                    LineItems(i).Quantity = IIf(Me.LineItems.Rows(i).Item("quantity") Is DBNull.Value, 1, Me.LineItems.Rows(i).Item("quantity"))
                    LineItems(i).Amount = IIf(Me.LineItems.Rows(i).Item("unit") Is DBNull.Value, 0, Me.LineItems.Rows(i).Item("unit"))
                    LineItems(i).Taxable = IIf(Me.LineItems.Rows(i).Item("taxable") Is DBNull.Value, True, Me.LineItems.Rows(i).Item("taxable"))
                Next
                Return Me.TaxCode.CalculateTax(LineItems, Me.Freight)
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property Subtotal() As Double
        Get
            Dim total As Double = 0
            For Each r As DataRow In Me.LineItems.Rows
                Try
                    total += r.Item("quantity") * r.Item("unit")
                Catch
                End Try
            Next
            Return total
        End Get
    End Property

    Public ReadOnly Property Total() As Double
        Get
            Return Me.Subtotal + Me.Freight + Me.SalesTax
        End Get
    End Property

    Public ReadOnly Property TravelTotal() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                If Row.Item("item_type_id") = 3 Then
                    Amt += Row.Item("quantity") * Row.Item("unit")
                End If
            Next
            Return Amt
        End Get
    End Property

    Public ReadOnly Property LaborTotal() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                If Row.Item("item_type_id") = 2 Then
                    Amt += Row.Item("quantity") * Row.Item("unit")
                End If
            Next
            Return Amt
        End Get
    End Property

    Public ReadOnly Property ItemsTotal() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                If Row.Item("item_type_id") = 1 Then
                    Amt += Row.Item("quantity") * Row.Item("unit")
                End If
            Next
            Return Amt
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
            Sql &= " AND department=" & CInt(cInteraction.ReferenceTypes.Sales)
            Sql &= " ORDER BY touch_date DESC"
            Return Me.Database.GetAll(Sql)
        End Get
    End Property

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
        Me.SetupLineItemsTable()
        Me.PopulateOffice()
        Me.PopulateShipVia()
        Me.PopulateEmployees()
        Me.PopulateTaxCodes()
        Me.TaxCode = New cTaxCode(db)
    End Sub

    Private Sub SetupLineItemsTable()
        Me.LineItems = New DataTable
        Me.LineItems.Columns.Add("quantity")
        Me.LineItems.Columns.Add("part_no")
        Me.LineItems.Columns.Add("serial_no")
        Me.LineItems.Columns.Add("equipment_id")
        Me.LineItems.Columns.Add("description")
        Me.LineItems.Columns.Add("unit")
        Me.LineItems.Columns.Add("ext_price")
        Me.LineItems.Columns.Add("station_id")
        Me.LineItems.Columns.Add("prime")
        Me.LineItems.Columns.Add("tax_status_id")
        Me.LineItems.Columns.Add("taxable")
        Me.LineItems.Columns.Add("item_type_id")
    End Sub

    Private Sub PopulateOffice()
        Me.Offices = Me.Database.GetAll("SELECT id, number, name, sort FROM office ORDER BY sort")
    End Sub

    Private Sub PopulateShipVia()
        Me.ShipVias = Me.Database.GetAll("SELECT id, name, sort FROM ship_via ORDER BY sort")
    End Sub

    Private Sub PopulateEmployees()
        Me.Employees = Me.Database.GetAll("SELECT windows_user, last_name + ', ' + first_name AS display_name FROM employee WHERE deactivated=0 ORDER BY last_name, first_name")
    End Sub

    Private Sub PopulateTaxCodes()
        Me.TaxCodes = Me.Database.GetAll("SELECT id, name FROM tax_code ORDER BY name")
    End Sub

    Public Function TaxStatuses() As DataTable
        Return Me.Database.GetAll("SELECT id, code, description, taxable FROM tax_status ORDER BY code")
    End Function

    Public Sub Open(ByVal Id As Integer)
        Dim Sql As String = ""
        Sql &= " SELECT so.*, "
        Sql &= " s.cst_name AS ship_to_name, s.cst_city + ', ' + s.cst_state AS ship_to_city,"
        Sql &= " b.cst_name AS bill_to_name, b.cst_city + ', ' + b.cst_state AS bill_to_city"
        Sql &= " FROM sales_order so"
        Sql &= " LEFT JOIN ADDRESS s ON so.ship_to=s.cst_no"
        Sql &= " LEFT JOIN ADDRESS b ON so.bill_to=b.cst_no"
        Sql &= " WHERE so.id=" & Id
        Dim r As DataRow = Me.Database.GetRow(Sql)
        If Me.Database.LastQuery.RowsReturned = 1 Then
            If Not r Is Nothing Then
                Me.BillTo = r.Item("bill_to")
                Me.Contact = r.Item("contact")
                Me.ContactID = r.Item("contact_id")
                Me.CreatedBy = r.Item("created_by")
                Me.DateCreated = r.Item("date_created")
                Me.DateDue = r.Item("date_due")
                Me.DateLastUpdated = r.Item("date_last_updated")
                Me.DateDelivered = IIf(r.Item("date_delivered") Is DBNull.Value, Nothing, r.Item("date_delivered"))
                Me.Fob = r.Item("fob")
                Me.LastUpdatedBy = r.Item("last_updated_by")
                Me.LeadId = r.Item("lead_id")
                Me.Notes = r.Item("notes")
                Me.Office = r.Item("office")
                Me.ProjectId = r.Item("project_id")
                Me.CustomerPO = r.Item("po")
                Me.QuoteId = r.Item("quote_id")
                Me.Salesperson = r.Item("salesperson")
                Me.Freight = r.Item("shipping_charge")
                Me.ShipTo = r.Item("ship_to")
                Me.ShipVia = r.Item("ship_via_id")
                Me.TaxCodeId = r.Item("tax_code_id")
                Me.Technician = r.Item("technician").ToString.Trim
                Me.TermsId = r.Item("terms_id")
                Me.TrackingNo = r.Item("tracking_no")
                Me.InvoiceNo = IIf(r.Item("invoice_no") Is DBNull.Value, "", r.Item("invoice_no"))
                Me.ServiceOrderId = r.Item("service_order_id")
                Me.InternalNotes = r.Item("internal_notes")
                Me.NotifyStateStatus = r.Item("notify_state_status")
                Me.Voided = r.Item("voided")
                Me._Id = Id
                ' Get line items
                Sql = "SELECT li.id, li.quantity, li.part_no, li.serial_no, li.equipment_id, li.description, li.unit,"
                Sql &= " (li.unit * li.quantity) AS ext_price,"
                Sql &= " li.tax_status_id, li.station_id, pn.item_type_id, pn.prime, ts.taxable"
                Sql &= " FROM sales_order_item li"
                Sql &= " LEFT OUTER JOIN item_master pn ON pn.part_no=li.part_no"
                Sql &= " LEFT OUTER JOIN tax_status ts ON ts.id=li.tax_status_id"
                Sql &= " WHERE li.sales_order_id=" & Id
                Me.LineItems = Me.Database.GetAll(Sql)
                ' Loop through them and 
                ' Return
                RaiseEvent Reload()
            Else
                Throw New Exception("Error opening sales order.")
            End If
        Else
            Throw New Exception("No sales order found with that number.")
        End If
    End Sub

    Public Sub Save()
        Dim Sql As String = ""
        If Me._Id > 0 Then
            ' Edit exiting sales order
            Sql &= "UPDATE sales_order"
            Sql &= " SET"
            Sql &= " bill_to=@bill_to, "
            Sql &= " ship_to=@ship_to, "
            Sql &= " project_id=@project_id, "
            Sql &= " salesperson=@salesperson, "
            Sql &= " po=@po, "
            Sql &= " date_due=@date_due, "
            Sql &= " date_delivered=@date_delivered,"
            Sql &= " ship_via_id=@ship_via_id, "
            Sql &= " tracking_no=@tracking_no, "
            Sql &= " quote_id=@quote_id, "
            Sql &= " contact=@contact_name, "
            Sql &= " contact_id=@contact_id, "
            Sql &= " tax_code_id=@tax_code_id, "
            Sql &= " office=@office, "
            Sql &= " notes=@notes, "
            Sql &= " technician=@technician, "
            Sql &= " fob=@fob, "
            Sql &= " shipping_charge=@shipping_charge,  "
            Sql &= " last_updated_by=@user,"
            Sql &= " date_last_updated=" & Me.Database.Timestamp & ","
            Sql &= " invoice_no=@invoice_no,"
            Sql &= " terms_id=@terms_id,"
            Sql &= " service_order_id=@service_order_id,"
            Sql &= " internal_notes=@internal_notes,"
            Sql &= " notify_state_status=@notify_state_status,"
            Sql &= " voided=@voided"
            Sql &= " WHERE id=@id"
            Sql = Sql.Replace("@bill_to", Me.Database.Escape(Me.BillTo))
            Sql = Sql.Replace("@ship_to", Me.Database.Escape(Me.ShipTo))
            Sql = Sql.Replace("@project_id", Me.ProjectId)
            Sql = Sql.Replace("@salesperson", Me.Database.Escape(Me.Salesperson))
            Sql = Sql.Replace("@po", Me.Database.Escape(Me.CustomerPO))
            Sql = Sql.Replace("@date_due", Me.Database.Escape(Me.DateDue))
            Sql = Sql.Replace("@date_delivered", Me.Database.Escape(IIf(Me.DateDelivered = Nothing, DBNull.Value, Me.DateDelivered)))
            Sql = Sql.Replace("@ship_via_id", Me.ShipVia)
            Sql = Sql.Replace("@tracking_no", Me.Database.Escape(Me.TrackingNo))
            Sql = Sql.Replace("@quote_id", Me.QuoteId)
            Sql = Sql.Replace("@lead_id", Me.LeadId)
            Sql = Sql.Replace("@contact_id", Me.Database.Escape(Me.ContactID))
            Sql = Sql.Replace("@contact_name", Me.Database.Escape(Me.Contact))
            Sql = Sql.Replace("@tax_code_id", Me.TaxCodeId)
            Sql = Sql.Replace("@office", Me.Database.Escape(Me.Office))
            Sql = Sql.Replace("@notes", Me.Database.Escape(Me.Notes))
            Sql = Sql.Replace("@technician", Me.Database.Escape(Me.Technician))
            Sql = Sql.Replace("@fob", Me.Database.Escape(Me.Fob))
            Sql = Sql.Replace("@shipping_charge", Me.Freight)
            Sql = Sql.Replace("@user", Me.Database.Escape(Me.LastUpdatedBy))
            Sql = Sql.Replace("@invoice_no", Me.Database.Escape(Me.InvoiceNo))
            Sql = Sql.Replace("@terms_id", Me.TermsId)
            Sql = Sql.Replace("@service_order_id", Me.ServiceOrderId)
            Sql = Sql.Replace("@internal_notes", Me.Database.Escape(Me.InternalNotes))
            Sql = Sql.Replace("@notify_state_status", Me.Database.Escape(Me.NotifyStateStatus))
            Sql = Sql.Replace("@voided", Me.Database.Escape(Me.Voided))
            Sql = Sql.Replace("@id", Me.OrderNo)
            Me.Database.Execute(Sql)
        Else
            ' Save new sales order
            Sql &= "INSERT INTO sales_order (bill_to, ship_to, project_id, salesperson, po, date_due, date_delivered,"
            Sql &= " ship_via_id, tracking_no, quote_id, contact, contact_id, tax_code_id, office, notes, notify_state_status,"
            Sql &= " technician, fob, shipping_charge, created_by, last_updated_by, voided,"
            Sql &= " date_last_updated, date_created, lead_id, invoice_no, terms_id, internal_notes, service_order_id)"
            Sql &= " VALUES ("
            Sql &= " @bill_to, @ship_to, @project_id, @salesperson, @po, @date_due, @date_delivered, @ship_via_id,"
            Sql &= " @tracking_no, @quote_id, @contact_name, @contact_id, @tax_code_id, @office, @notes, @notify_state_status,"
            Sql &= " @technician, @fob, @shipping_charge, @user, @user, @voided,"
            Sql &= " " & Me.Database.Timestamp & ", " & Me.Database.Timestamp & ", @lead_id, @invoice_no, @terms_id, @internal_notes, @service_order_id"
            Sql &= " )"
            Sql = Sql.Replace("@bill_to", Me.Database.Escape(Me.BillTo))
            Sql = Sql.Replace("@ship_to", Me.Database.Escape(Me.ShipTo))
            Sql = Sql.Replace("@project_id", Me.ProjectId)
            Sql = Sql.Replace("@salesperson", Me.Database.Escape(Me.Salesperson))
            Sql = Sql.Replace("@po", Me.Database.Escape(Me.CustomerPO))
            Sql = Sql.Replace("@date_due", Me.Database.Escape(Me.DateDue))
            Sql = Sql.Replace("@date_delivered", Me.Database.Escape(IIf(Me.DateDelivered = Nothing, DBNull.Value, Me.DateDelivered)))
            Sql = Sql.Replace("@ship_via_id", Me.ShipVia)
            Sql = Sql.Replace("@tracking_no", Me.Database.Escape(Me.TrackingNo))
            Sql = Sql.Replace("@quote_id", Me.QuoteId)
            Sql = Sql.Replace("@lead_id", Me.LeadId)
            Sql = Sql.Replace("@contact_name", Me.Database.Escape(Me.Contact))
            Sql = Sql.Replace("@contact_id", Me.Database.Escape(Me.ContactID))
            Sql = Sql.Replace("@tax_code_id", Me.TaxCodeId)
            Sql = Sql.Replace("@office", Me.Database.Escape(Me.Office))
            Sql = Sql.Replace("@notes", Me.Database.Escape(Me.Notes))
            Sql = Sql.Replace("@technician", Me.Database.Escape(Me.Technician))
            Sql = Sql.Replace("@fob", Me.Database.Escape(Me.Fob))
            Sql = Sql.Replace("@shipping_charge", Me.Freight)
            Sql = Sql.Replace("@user", Me.Database.Escape(Me.LastUpdatedBy))
            Sql = Sql.Replace("@invoice_no", Me.Database.Escape(Me.InvoiceNo))
            Sql = Sql.Replace("@terms_id", Me.TermsId)
            Sql = Sql.Replace("@service_order_id", Me.ServiceOrderId)
            Sql = Sql.Replace("@internal_notes", Me.Database.Escape(Me.InternalNotes))
            Sql = Sql.Replace("@notify_state_status", Me.Database.Escape(Me.NotifyStateStatus))
            Sql = Sql.Replace("@voided", Me.Database.Escape(Me.Voided))
            Me.Database.InsertAndReturnId(Sql)
        End If
        If Me.Database.LastQuery.Successful Then
            ' If new, set id
            If Me._Id = 0 Then
                Me._Id = Me.Database.LastQuery.InsertId
            End If
            RaiseEvent Saved(Me)
            ' Save Line Items
            Dim Row As DataRow
            If Me.LineItems.Rows.Count > 0 Then
                For Each Row In Me.LineItems.Rows
                    If Row.RowState = DataRowState.Added Then
                        Dim qty As Double = IIf(Row.Item("quantity").ToString = "", 0, Row.Item("quantity"))
                        If qty > 0 Then
                            Sql = "INSERT INTO sales_order_item"
                            Sql &= " (sales_order_id, quantity, part_no, serial_no, equipment_id, station_id, "
                            Sql &= " description, unit, date_last_updated, date_created, tax_status_id)"
                            Sql &= " VALUES"
                            Sql &= " (@sales_order_id, @quantity, @part_no, @serial_no, @equipment_id,"
                            Sql &= " @station_id, @description, @unit, " & Me.Database.Timestamp & ", " & Me.Database.Timestamp & ", @tax_status_id)"
                            Sql = Sql.Replace("@sales_order_id", Me._Id)
                            Sql = Sql.Replace("@quantity", qty)
                            Sql = Sql.Replace("@part_no", Me.Database.Escape(Row.Item("part_no")))
                            Sql = Sql.Replace("@serial_no", Me.Database.Escape(IIf(Row.Item("serial_no") Is DBNull.Value, "", Row.Item("serial_no"))))
                            Sql = Sql.Replace("@equipment_id", IIf(Row.Item("equipment_id") Is DBNull.Value, 0, Row.Item("equipment_id")))
                            Sql = Sql.Replace("@description", Me.Database.Escape(IIf(Row.Item("description") Is DBNull.Value, "", Row.Item("description"))))
                            Sql = Sql.Replace("@unit", IIf(Row.Item("unit").ToString = "", 0, Row.Item("unit")))
                            Sql = Sql.Replace("@station_id", IIf(Row.Item("station_id").ToString = "", 0, Row.Item("station_id")))
                            Sql = Sql.Replace("@tax_status_id", Row.Item("tax_status_id"))
                            Me.Database.Execute(Sql)
                            If Not Me.Database.LastQuery.Successful Then
                                Dim Err As String = Me.Database.LastQuery.ErrorMsg
                            End If
                        End If
                    ElseIf Row.RowState = DataRowState.Modified Then
                        Dim qty As Double = IIf(Row.Item("quantity").ToString = "", 0, Row.Item("quantity"))
                        If qty > 0 Then
                            Sql = "UPDATE sales_order_item SET"
                            Sql &= " sales_order_id=@sales_order_id,"
                            Sql &= " quantity=@quantity, "
                            Sql &= " part_no=@part_no,     "
                            Sql &= " station_id=@station_id,"
                            Sql &= " serial_no=@serial_no, "
                            Sql &= " equipment_id=@equipment_id,"
                            Sql &= " description=@description, "
                            Sql &= " unit=@unit, "
                            Sql &= " tax_status_id=@tax_status_id"
                            Sql &= " WHERE id=@id"
                            Sql = Sql.Replace("@sales_order_id", Me._Id)
                            Sql = Sql.Replace("@quantity", qty)
                            Sql = Sql.Replace("@part_no", Me.Database.Escape(Row.Item("part_no")))
                            Sql = Sql.Replace("@equipment_id", IIf(Row.Item("equipment_id") Is DBNull.Value, 0, Row.Item("equipment_id")))
                            Sql = Sql.Replace("@serial_no", Me.Database.Escape(IIf(Row.Item("serial_no") Is DBNull.Value, "", Row.Item("serial_no"))))
                            Sql = Sql.Replace("@description", Me.Database.Escape(IIf(Row.Item("description") Is DBNull.Value, "", Row.Item("description"))))
                            Sql = Sql.Replace("@unit", IIf(Row.Item("unit").ToString = "", 0, Row.Item("unit")))
                            Sql = Sql.Replace("@station_id", IIf(Row.Item("station_id").ToString = "", 0, Row.Item("station_id")))
                            Sql = Sql.Replace("@tax_status_id", Row.Item("tax_status_id"))
                            Sql = Sql.Replace("@id", Row.Item("id"))
                            Me.Database.Execute(Sql)
                            If Not Me.Database.LastQuery.Successful Then
                                MsgBox(Me.Database.LastQuery.ErrorMsg)
                            End If
                        Else
                            ' Delete quantity 0s
                            Sql = "DELETE FROM sales_order_item WHERE id=" & Row.Item("id")
                            Me.Database.Execute(Sql)
                        End If
                    End If
                Next
            End If
            Me.Open(Me._Id)
        Else
            Throw New Exception(Me.Database.LastQuery.ErrorMsg)
        End If
    End Sub


    Public Function GetPurchaseOrderrs() As DataTable
        Dim Sql As String = ""
        Sql &= "SELECT po.po_no, po.vendor_no, c.cst_name AS vendor_name, "
        Sql &= " po.po_date, po.date_ordered, po.date_planned_ship, po.date_expected,"
        Sql &= " ((SELECT SUM(quantity*unit_price) FROM purchase_order_item WHERE po_no=po.po_no) + shipping_charge + tax) AS total_price,"
        Sql &= " po.office, po.requested_by,"
        Sql &= " type = CASE po.rma WHEN 1 THEN 'RMA' ELSE 'PO' END"
        Sql &= " FROM purchase_order po, ADDRESS c"
        Sql &= " WHERE type=2 AND our_order_no=" & Me.Database.Escape(Me.OrderNo)
        Sql &= " ORDER BY po_date"
        Return Me.Database.GetAll(Sql)
    End Function

    Public Function DeleteLineItem(ByVal Id As Integer) As Boolean
        Me.Database.Execute("DELTE FROM sales_order_item WHERE id=" & Id)
        Return Me.Database.LastQuery.Successful
    End Function

    Public Function ToGravityDocument(ByVal Template As String) As GravityDocument.gDocument
        ' If no template specified
        If Template.Length = 0 Then
            Dim id As Integer = Me.Database.GetOne("SELECT value FROM settings WHERE property='Template Sales Order'")
            Template = Me.Database.GetOne("SELECT html FROM template WHERE id=" & id)
        End If
        ' Create Gravity Document
        Dim Doc As New GravityDocument.gDocument(Me.Database.GetOne("SELECT value FROM settings WHERE property='Page Height in Pixels'"))
        Doc.LoadXml(Template)
        ' Settings
        Doc.FormType = GravityDocument.gDocument.FormTypes.SalesOrder
        Doc.ReferenceID = Me.OrderNo
        ' Put in variables
        ' Get Ship To
        Dim Sql As String = "SELECT cst_name, cst_addr1, cst_addr2, cst_city, cst_state, cst_zip, cst_phone, cst_fax"
        Sql &= " FROM ADDRESS WHERE cst_no=" & Me.Database.Escape(Me.ShipTo)
        Dim ShipTo As DataRow = Me.Database.GetRow(Sql)
        ' Create address strings
        Dim ShipToAddress As String = ""
        ' Set Ship To
        ShipToAddress = ShipTo.Item("cst_addr1")
        If Me.IfNull(ShipTo.Item("cst_addr2")).Length > 0 Then
            ShipToAddress &= ControlChars.CrLf & ShipTo.Item("cst_addr2")
        End If
        ' Replace variables
        Dim Page As GravityDocument.gPage = Doc.GetPage(1)
        Page.AddVariable("%order_no%", Me.OrderNo)
        Page.AddVariable("%due_date%", Format(Me.DateDue, "MM/dd/yy"))
        Page.AddVariable("%created_date%", Format(Me.DateCreated, "MM/dd/yy"))
        Page.AddVariable("%delivered_date%", Format(Me.DateDelivered, "MM/dd/yy"))
        Page.AddVariable("%ship_to_name%", Me.IfNull(ShipTo.Item("cst_name")))
        Page.AddVariable("%ship_to_address%", ShipToAddress)
        Page.AddVariable("%ship_to_city%", Me.IfNull(ShipTo.Item("cst_city")))
        Page.AddVariable("%ship_to_state%", Me.IfNull(ShipTo.Item("cst_state")))
        Page.AddVariable("%ship_to_zip%", Me.IfNull(ShipTo.Item("cst_zip")))
        Page.AddVariable("%ship_to_phone%", Me.IfNull(ShipTo.Item("cst_phone")))
        Page.AddVariable("%ship_to_fax%", Me.IfNull(ShipTo.Item("cst_fax")))
        Dim BillToCo As New cCompany(Me.Database)
        BillToCo.Open(Me.BillTo)
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
        Page.AddVariable("%bill_to_no%", Me.BillTo)
        Page.AddVariable("%ship_to_no%", Me.ShipTo)
        Page.AddVariable("%po%", Me.CustomerPO)
        Page.AddVariable("%contact%", Me.Contact)
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
        ' Order Details
        If Me.DateDelivered = Nothing Then
            Page.AddVariable("%delivered%", "--")
        Else
            Page.AddVariable("%delivered%", Format(Me.DateDelivered, "MM/dd/yy"))
        End If
        Page.AddVariable("%fob%", Me.Fob)
        Page.AddVariable("%ship_via%", Me.Database.GetOne("SELECT name FROM ship_via WHERE id=" & Me.ShipVia))
        Page.AddVariable("%salesperson%", Me.Salesperson)
        Page.AddVariable("%notes%", Me.Notes)
        ' Billing
        Page.AddVariable("%subtotal%", Format(Me.Subtotal, "$0.00"))
        Page.AddVariable("%freight%", Format(Me.Freight, "$0.00"))
        Page.AddVariable("%tax%", Format(Me.SalesTax, "$0.00"))
        Page.AddVariable("%total%", Format(Me.Total, "$0.00"))
        ' Line Items
        Dim Settings As New cSettings(Me.Database)
        Dim GroupSamePart As Integer = Settings.GetValue("Sales Order Part Grouping", 1)
        Dim Table As New DataTable
        Dim Row As DataRow
        Table.Columns.Add("quantity")
        Table.Columns.Add("part_no")
        Table.Columns.Add("serial_no")
        Table.Columns.Add("description")
        Table.Columns.Add("unit_price")
        Table.Columns.Add("ext_price")
        Try
            If GroupSamePart = 1 Then
                Dim Hash As New Hashtable
                Dim Index As Integer = 0
                For Each Row In Me.LineItems.Rows
                    Dim Key As String = Row.Item("part_no") & "__" & Row.Item("unit").ToString.Replace(".", "_")
                    If Hash.ContainsKey(Key) Then
                        Table.Rows(Hash.Item(Key)).Item("quantity") += Row.Item("quantity")
                        Table.Rows(Hash.Item(Key)).Item("serial_no") = ""
                        Table.Rows(Hash.Item(Key)).Item("ext_price") += Row.Item("quantity") * Row.Item("unit")
                    Else
                        Hash.Add(Key, Index)
                        Dim r As DataRow = Table.NewRow
                        r.Item("quantity") = Row.Item("quantity")
                        r.Item("part_no") = Row.Item("part_no")
                        r.Item("serial_no") = Row.Item("serial_no")
                        r.Item("description") = Row.Item("description")
                        r.Item("unit_price") = Row.Item("unit")
                        r.Item("ext_price") = Row.Item("unit") * Row.Item("quantity")
                        Table.Rows.Add(r)
                        Index += 1
                    End If
                Next
            Else
                For Each Row In Me.LineItems.Rows
                    Dim r As DataRow = Table.NewRow
                    r.Item("quantity") = Row.Item("quantity")
                    r.Item("part_no") = Row.Item("part_no")
                    r.Item("serial_no") = Row.Item("serial_no")
                    r.Item("description") = Row.Item("description")
                    r.Item("unit_price") = Row.Item("unit")
                    r.Item("ext_price") = Row.Item("unit") * Row.Item("quantity")
                    Table.Rows.Add(r)
                Next
            End If
            Dim Element As GravityDocument.gElement = Page.GetTableBySource("line_items")
            If Element IsNot Nothing Then
                Element.Table.Data = Table
            End If
        Catch ex As Exception
            Throw New Exception("Error generating line items. " & ex.ToString)
        End Try
        
        ' Return gDocument
        Return Doc
    End Function

    Private Function IfNull(ByVal Value As Object, Optional ByVal DefaultVal As Object = "") As Object
        If Value Is DBNull.Value Then
            Return DefaultVal
        Else
            Return Value
        End If
    End Function


End Class
