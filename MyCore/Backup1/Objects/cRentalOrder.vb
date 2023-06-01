Imports MyCore.Data

Public Class cRentalOrder

    Dim _Id As Integer = 0

    Public BillTo As String = ""
    Public ShipTo As String = ""
    Public ProjectId As Integer = 0
    Public Salesperson As String = ""
    Public CustomerPO As String = ""
    Public DateCreated As Date = Now
    Public DateLastUpdated As Date = Now
    Public DateOrdered As Date = Now
    Public DateStart As Date = Now
    Public DateDelivered As Date = Nothing
    Public DateReturned As Date = Nothing
    Public NumberOfDays As Integer = 7
    Public CustomerPickup As Boolean = False
    Public QuoteId As Integer = 0
    Public LeadId As Integer = 0
    Public Contact As String = ""
    Public ContactID As Integer = 0
    Dim _TaxCodeId As Integer = 0
    Public Office As Integer = 0
    Public Notes As String = ""
    Public InternalNotes As String = ""
    Public CreatedBy As String = ""
    Public LastUpdatedBy As String = ""
    Public Technician As String = ""
    Public InvoiceNo As String = ""
    Public TermsId As Integer = 0
    Public Voided As Boolean = False

    Public LineItems As DataTable
    Public Offices As DataTable
    Public Employees As DataTable
    Public TaxCodes As DataTable

    Public TaxCode As cTaxCode

    Public Database As MyCore.Data.EasySql
    Public Event Reload()
    Public Event Saved(ByVal RentalOrder As cRentalOrder)

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

    Public ReadOnly Property DateExpected() As Date
        Get
            Return Me.DateStart.AddDays(Me.NumberOfDays)
        End Get
    End Property

    Public ReadOnly Property OrderNo() As Integer
        Get
            Return Me._Id
        End Get
    End Property

    Public ReadOnly Property LineItemsForSalesTaxCalculation() As cTaxCode.LineItem()
        Get
            ' Line Items
            Dim LineItems(Me.LineItems.Rows.Count - 1) As cTaxCode.LineItem
            For i As Integer = 0 To Me.LineItems.Rows.Count - 1
                LineItems(i).Quantity = Me.IfNull(Me.LineItems.Rows(i).Item("quantity"), 1)
                LineItems(i).Amount = Me.IfNull(Me.LineItems.Rows(i).Item("unit"), 0)
                LineItems(i).Taxable = Me.IfNull(Me.LineItems.Rows(i).Item("taxable"), 0)
            Next
            Return LineItems
        End Get
    End Property

    Public ReadOnly Property SalesTax() As Double
        Get
            If Me.TaxCodeId > 0 Then                
                Return Me.TaxCode.CalculateTax(Me.LineItemsForSalesTaxCalculation, 0)
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property SalesTaxLineItems() As DataTable
        Get
            If Me.TaxCodeId > 0 Then
                Dim TaxCode As New cTaxCode(Me.Database)
                TaxCode.Open(Me.TaxCodeId)
                Return TaxCode.TaxPerAuthority(Me.LineItemsForSalesTaxCalculation, 0)
            Else
                Return Nothing
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

    Public ReadOnly Property ServiceOrders() As DataTable
        Get
            Dim Sql As String = "SELECT so.id, date_completed, created_by, date_created"
            Sql &= ", ship_to.cst_name AS ship_to_name, ship_to.cst_no AS ship_to_no"
            Sql &= " FROM service_order so"
            Sql &= " JOIN ADDRESS ship_to ON so.location_id=ship_to.cst_no"
            Sql &= " WHERE so.invoice_id=" & Me.OrderNo
            Sql &= " AND so.charge_to=2"
            If Me.Database.LastQuery.Successful Then
                Return Me.Database.GetAll(Sql)
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
            Sql &= " AND department=" & CInt(cInteraction.ReferenceTypes.Rental)
            Return Me.Database.GetAll(Sql)
        End Get
    End Property

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
        Me.SetupLineItemsTable()
        Me.PopulateOffice()
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

    Private Sub PopulateEmployees()
        Me.Employees = Me.Database.GetAll("SELECT windows_user, last_name + ', ' + first_name AS display_name FROM employee WHERE deactivated=0 ORDER BY last_name, first_name")
    End Sub

    Private Sub PopulateTaxCodes()
        Me.TaxCodes = Me.Database.GetAll("SELECT id, name FROM tax_code ORDER BY name")
    End Sub

    Public Function TaxStatuses() As DataTable
        Return Me.Database.GetAll("SELECT id, code, description, taxable FROM tax_status ORDER BY code")
    End Function

    Public Function Open(ByVal Id As Integer) As Boolean
        Dim Sql As String = ""
        Sql &= "SELECT ro.*, "
        Sql &= " s.cst_name AS ship_to_name, s.cst_city + ', ' + s.cst_state AS ship_to_city,"
        Sql &= " b.cst_name AS bill_to_name, b.cst_city + ', ' + b.cst_state AS bill_to_city"
        Sql &= " FROM rental_order ro"
        Sql &= " LEFT JOIN ADDRESS s ON ro.ship_to_no=s.cst_no"
        Sql &= " LEFT JOIN ADDRESS b ON ro.bill_to_no=b.cst_no"
        Sql &= " WHERE ro.id=" & Id
        Dim r As DataRow = Me.Database.GetRow(Sql)
        If Me.Database.LastQuery.RowsReturned = 1 Then
            If r IsNot Nothing Then
                Me.BillTo = r.Item("bill_to_no")
                Me.Contact = r.Item("contact")
                Me.ContactID = r.Item("contact_id")
                Me.CreatedBy = r.Item("created_by")
                Me.DateCreated = r.Item("date_created")
                Me.DateOrdered = r.Item("date_ordered")
                Me.DateLastUpdated = r.Item("date_last_updated")
                Me.DateDelivered = IIf(r.Item("date_delivered") Is DBNull.Value, Nothing, r.Item("date_delivered"))
                Me.DateReturned = IIf(r.Item("date_returned") Is DBNull.Value, Nothing, r.Item("date_returned"))
                Me.DateStart = r.Item("date_start")
                Me.LastUpdatedBy = r.Item("last_updated_by")
                Me.LeadId = r.Item("lead_id")
                Me.Notes = r.Item("notes")
                Me.InternalNotes = r.Item("internal_notes")
                Me.Office = r.Item("office")
                Me.ProjectId = r.Item("project_id")
                Me.CustomerPO = r.Item("po_no")
                Me.QuoteId = r.Item("quote_id")
                Me.Salesperson = r.Item("salesperson")
                Me.ShipTo = r.Item("ship_to_no")
                Me.TaxCodeId = r.Item("tax_code_id")
                Me.Technician = r.Item("technician").ToString.Trim
                Me.NumberOfDays = r.Item("duration_days")
                Me.CustomerPickup = r.Item("customer_pickup")
                Me.TermsId = r.Item("terms_id")
                Me.InvoiceNo = r.Item("invoice_no")
                Me.Voided = r.Item("voided")
                Me._Id = Id
                ' Get line items
                Sql = "SELECT li.id, li.quantity, li.part_no, li.serial_no, li.equipment_id, li.description, li.unit,"
                Sql &= " (li.unit * li.quantity) AS ext_price,"
                Sql &= " li.tax_status_id, li.station_id, pn.item_type_id, pn.prime, ts.taxable"
                Sql &= " FROM rental_order_item li"
                Sql &= " LEFT OUTER JOIN item_master pn ON pn.part_no=li.part_no"
                Sql &= " LEFT OUTER JOIN tax_status ts ON ts.id=li.tax_status_id"
                Sql &= " WHERE li.rental_order_id=" & Id
                Me.LineItems = Me.Database.GetAll(Sql)
                ' Return
                RaiseEvent Reload()
                Return True
            Else
                Return False
            End If
        Else
            Throw New Exception("No rental order found with that number.")
            Return False
        End If
    End Function

    Public Function Save() As Boolean
        Dim Sql As String = ""
        If Me._Id > 0 Then
            ' Edit exiting rental order
            Sql &= "UPDATE rental_order"
            Sql &= " SET"
            Sql &= " bill_to_no=@bill_to, "
            Sql &= " ship_to_no=@ship_to, "
            Sql &= " contact=@contact_name, "
            Sql &= " contact_id=@contact_id, "
            Sql &= " po_no=@po, "
            Sql &= " date_start=@date_start, "
            Sql &= " duration_days=@duration_days, "
            Sql &= " date_delivered=@date_delivered, "
            Sql &= " date_returned=@date_returned, "
            Sql &= " customer_pickup=@customer_pickup,"
            Sql &= " notes=@notes, "
            Sql &= " internal_notes=@internal_notes, "
            Sql &= " project_id=@project_id, "
            Sql &= " quote_id=@quote_id, "
            Sql &= " lead_id=@lead_id, "
            Sql &= " salesperson=@salesperson, "
            Sql &= " technician=@technician, "
            Sql &= " office=@office, "
            Sql &= " tax_code_id=@tax_code_id, "
            Sql &= " invoice_no=@invoice_no, "
            Sql &= " date_ordered=@date_ordered, "
            Sql &= " terms_id=@terms, "
            Sql &= " voided=@voided, "
            Sql &= " date_last_updated=" & Me.Database.Timestamp & ", "
            Sql &= " last_updated_by=@user"
            Sql &= " WHERE id=@id"
            Sql = Sql.Replace("@bill_to", Me.Database.Escape(Me.BillTo))
            Sql = Sql.Replace("@ship_to", Me.Database.Escape(Me.ShipTo))
            Sql = Sql.Replace("@contact_name", Me.Database.Escape(Me.Contact))
            Sql = Sql.Replace("@contact_id", Me.Database.Escape(Me.ContactID))
            Sql = Sql.Replace("@po", Me.Database.Escape(Me.CustomerPO))
            Sql = Sql.Replace("@date_start", Me.Database.Escape(Me.DateStart))
            Sql = Sql.Replace("@duration_days", Me.NumberOfDays)
            Sql = Sql.Replace("@date_delivered", Me.Database.Escape(IIf(Me.DateDelivered = Nothing, DBNull.Value, Me.DateDelivered)))
            Sql = Sql.Replace("@date_returned", Me.Database.Escape(IIf(Me.DateReturned = Nothing, DBNull.Value, Me.DateReturned)))
            Sql = Sql.Replace("@customer_pickup", Me.Database.Escape(Me.CustomerPickup))
            Sql = Sql.Replace("@notes", Me.Database.Escape(Me.Notes))
            Sql = Sql.Replace("@internal_notes", Me.Database.Escape(Me.InternalNotes))
            Sql = Sql.Replace("@project_id", Me.ProjectId)
            Sql = Sql.Replace("@quote_id", Me.QuoteId)
            Sql = Sql.Replace("@lead_id", Me.LeadId)
            Sql = Sql.Replace("@salesperson", Me.Database.Escape(Me.Salesperson))
            Sql = Sql.Replace("@technician", Me.Database.Escape(Me.Technician))
            Sql = Sql.Replace("@office", Me.Database.Escape(Me.Office))
            Sql = Sql.Replace("@tax_code_id", Me.TaxCodeId)
            Sql = Sql.Replace("@invoice_no", Me.Database.Escape(Me.InvoiceNo))
            Sql = Sql.Replace("@date_ordered", Me.Database.Escape(Me.DateOrdered))
            Sql = Sql.Replace("@user", Me.Database.Escape(Me.LastUpdatedBy))
            Sql = Sql.Replace("@terms", Me.TermsId)
            Sql = Sql.Replace("@voided", Me.Database.Escape(Me.Voided))
            Sql = Sql.Replace("@id", Me._Id)
            Me.Database.Execute(Sql)
        Else
            ' Save new rental order
            Sql = " INSERT INTO rental_order"
            Sql &= " (bill_to_no, ship_to_no, contact, contact_id, po_no, date_start, duration_days, date_delivered, date_returned, customer_pickup,"
            Sql &= " notes, internal_notes, project_id, quote_id, lead_id, salesperson, technician, office, tax_code_id, invoice_no, "
            Sql &= " date_ordered, terms_id, date_created, date_last_updated, created_by, last_updated_by, voided)"
            Sql &= " VALUES"
            Sql &= " (@bill_to, @ship_to, @contact_name, @contact_id, @po, @date_start, @duration_days, @date_delivered, @date_returned, @customer_pickup,"
            Sql &= " @notes, @internal_notes, @project_id, @quote_id, @lead_id, @salesperson, @technician, @office, @tax_code_id, @invoice_no,"
            Sql &= " @date_ordered, @terms, "
            Sql &= " " & Me.Database.Timestamp & ", " & Me.Database.Timestamp & ", @user, @user, @voided)"
            Sql = Sql.Replace("@bill_to", Me.Database.Escape(Me.BillTo))
            Sql = Sql.Replace("@ship_to", Me.Database.Escape(Me.ShipTo))
            Sql = Sql.Replace("@contact_name", Me.Database.Escape(Me.Contact))
            Sql = Sql.Replace("@contact_id", Me.Database.Escape(Me.ContactID))
            Sql = Sql.Replace("@po", Me.Database.Escape(Me.CustomerPO))
            Sql = Sql.Replace("@date_start", Me.Database.Escape(Me.DateStart))
            Sql = Sql.Replace("@duration_days", Me.NumberOfDays)
            Sql = Sql.Replace("@date_delivered", Me.Database.Escape(IIf(Me.DateDelivered = Nothing, DBNull.Value, Me.DateDelivered)))
            Sql = Sql.Replace("@date_returned", Me.Database.Escape(IIf(Me.DateReturned = Nothing, DBNull.Value, Me.DateReturned)))
            Sql = Sql.Replace("@customer_pickup", Me.Database.Escape(Me.CustomerPickup))
            Sql = Sql.Replace("@notes", Me.Database.Escape(Me.Notes))
            Sql = Sql.Replace("@internal_notes", Me.Database.Escape(Me.InternalNotes))
            Sql = Sql.Replace("@project_id", Me.ProjectId)
            Sql = Sql.Replace("@quote_id", Me.QuoteId)
            Sql = Sql.Replace("@lead_id", Me.LeadId)
            Sql = Sql.Replace("@salesperson", Me.Database.Escape(Me.Salesperson))
            Sql = Sql.Replace("@technician", Me.Database.Escape(Me.Technician))
            Sql = Sql.Replace("@office", Me.Database.Escape(Me.Office))
            Sql = Sql.Replace("@tax_code_id", Me.TaxCodeId)
            Sql = Sql.Replace("@invoice_no", Me.Database.Escape(Me.InvoiceNo))
            Sql = Sql.Replace("@date_ordered", Me.Database.Escape(Me.DateOrdered))
            Sql = Sql.Replace("@user", Me.Database.Escape(Me.LastUpdatedBy))
            Sql = Sql.Replace("@terms", Me.TermsId)
            Sql = Sql.Replace("@voided", Me.Database.Escape(Me.Voided))
            Me.Database.InsertAndReturnId(Sql)
            If Not Me.Database.LastQuery.Successful Then
                Dim Err As String = Me.Database.LastQuery.ErrorMsg
            End If
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
                        Sql = "INSERT INTO rental_order_item"
                        Sql &= " (rental_order_id, quantity, part_no, serial_no, equipment_id, station_id, "
                        Sql &= " description, unit, date_last_updated, tax_status_id)"
                        Sql &= " VALUES"
                        Sql &= " (@rental_order_id, @quantity, @part_no, @serial_no, @equipment_id,"
                        Sql &= " @station_id, @description, @unit, " & Me.Database.Timestamp & ", @tax_status_id)"
                        Sql = Sql.Replace("@rental_order_id", Me._Id)
                        Sql = Sql.Replace("@quantity", Row.Item("quantity"))
                        Sql = Sql.Replace("@part_no", Me.Database.Escape(Row.Item("part_no")))
                        Sql = Sql.Replace("@serial_no", Me.Database.Escape(IIf(Row.Item("serial_no") Is DBNull.Value, "", Row.Item("serial_no"))))
                        Sql = Sql.Replace("@equipment_id", IIf(Row.Item("equipment_id") Is DBNull.Value, 0, Row.Item("equipment_id")))
                        Sql = Sql.Replace("@description", Me.Database.Escape(IIf(Row.Item("description") Is DBNull.Value, "", Row.Item("description"))))
                        Sql = Sql.Replace("@unit", IIf(Row.Item("unit").ToString = "", 0, Row.Item("unit")))
                        Sql = Sql.Replace("@station_id", IIf(Row.Item("station_id").ToString = "", 0, Row.Item("station_id")))
                        Sql = Sql.Replace("@tax_status_id", Row.Item("tax_status_id"))
                        Me.Database.Execute(Sql)
                        If Not Me.Database.LastQuery.Successful Then
                            MsgBox(Me.Database.LastQuery.ErrorMsg)
                        End If
                    ElseIf Row.RowState = DataRowState.Modified Then
                        If Row.Item("quantity") > 0 Then
                            Sql = "UPDATE rental_order_item SET"
                            Sql &= " rental_order_id=@rental_order_id,"
                            Sql &= " quantity=@quantity, "
                            Sql &= " part_no=@part_no,     "
                            Sql &= " station_id=@station_id,"
                            Sql &= " serial_no=@serial_no, "
                            Sql &= " equipment_id=@equipment_id,"
                            Sql &= " description=@description, "
                            Sql &= " unit=@unit, "
                            Sql &= " tax_status_id=@tax_status_id"
                            Sql &= " WHERE id=@id"
                            Sql = Sql.Replace("@rental_order_id", Me._Id)
                            Sql = Sql.Replace("@quantity", Row.Item("quantity"))
                            Sql = Sql.Replace("@part_no", Me.Database.Escape(Row.Item("part_no")))
                            Sql = Sql.Replace("@serial_no", Me.Database.Escape(Row.Item("serial_no")))
                            Sql = Sql.Replace("@equipment_id", IIf(Row.Item("equipment_id") Is DBNull.Value, 0, Row.Item("equipment_id")))
                            Sql = Sql.Replace("@description", Me.Database.Escape(IIf(Row.Item("description") Is DBNull.Value, "", Row.Item("description"))))
                            Sql = Sql.Replace("@unit", IIf(Row.Item("unit").ToString = "", 0, Row.Item("unit")))
                            Sql = Sql.Replace("@station_id", IIf(Row.Item("station_id").ToString = "", 0, Row.Item("station_id")))
                            Sql = Sql.Replace("@tax_status_id", Row.Item("tax_status_id"))
                            Sql = Sql.Replace("@id", Row.Item("id"))
                            Me.Database.Execute(Sql)
                        Else
                            Me.DeleteLineItem(Row.Item("id"))
                        End If
                    End If
                Next
            End If
            Me.Open(Me._Id)
            Return True
        Else
            Throw New Exception(Me.Database.LastQuery.ErrorMsg)
            Return False
        End If
    End Function

    Public Function DeleteLineItem(ByVal Id As Integer)
        Me.Database.Execute("DELETE FROM rental_order_item WHERE id=" & Id)
        If Me.Database.LastQuery.Successful Then
            Return True
        Else
            Dim ErrorMsg As String = Me.Database.LastQuery.ErrorMsg
            Return False
        End If
    End Function

    Public Function ToGravityDocument(ByVal Template As String) As GravityDocument.gDocument
        ' If no template specified
        If Template.Length = 0 Then
            Dim id As Integer = Me.Database.GetOne("SELECT value FROM settings WHERE property='Template Rental Order'")
            Template = Me.Database.GetOne("SELECT html FROM template WHERE id=" & id)
        End If
        ' Create Gravity Document
        Dim Doc As New GravityDocument.gDocument(Me.Database.GetOne("SELECT value FROM settings WHERE property='Page Height in Pixels'"))
        Doc.LoadXml(Template)
        ' Settings
        Doc.FormType = GravityDocument.gDocument.FormTypes.RentalOrder
        Doc.ReferenceID = Me.OrderNo
        ' Put in variables
        ' Get Ship To
        Dim Sql As String = "SELECT cst_name, cst_addr1, cst_addr2, cst_city, cst_state, cst_zip"
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
        Page.AddVariable("%date_start%", Format(Me.DateStart, "MM/dd/yy"))
        Page.AddVariable("%date_expected%", Format(DateAdd(DateInterval.Day, Me.NumberOfDays, Me.DateStart), "MM/dd/yy"))
        Page.AddVariable("%date_ordered%", Format(Me.DateOrdered, "MM/dd/yy"))
        Page.AddVariable("%date_created%", Format(Me.DateCreated, "MM/dd/yy"))
        Page.AddVariable("%ship_to_name%", Me.IfNull(ShipTo.Item("cst_name")))
        Page.AddVariable("%ship_to_address%", ShipToAddress)
        Page.AddVariable("%ship_to_city%", Me.IfNull(ShipTo.Item("cst_city")))
        Page.AddVariable("%ship_to_state%", Me.IfNull(ShipTo.Item("cst_state")))
        Page.AddVariable("%ship_to_zip%", Me.IfNull(ShipTo.Item("cst_zip")))
        ' Bill to
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
        Page.AddVariable("%po%", Me.CustomerPO)
        If Me.DateDelivered = Nothing Then
            Page.AddVariable("%delivered%", "--")
        Else
            Page.AddVariable("%delivered%", Format(Me.DateDelivered, "MM/dd/yy"))
        End If
        If Me.DateReturned = Nothing Then
            Page.AddVariable("%date_returned%", "--")
        Else
            Page.AddVariable("%date_returned%", Format(Me.DateReturned, "MM/dd/yy"))
        End If
        Page.AddVariable("%contact%", Me.Contact)
        Page.AddVariable("%duration%", Me.NumberOfDays)
        Page.AddVariable("%salesperson%", Me.Salesperson)
        Page.AddVariable("%notes%", Me.Notes)
        Page.AddVariable("%technician%", Me.Technician)
        Page.AddVariable("%terms%", Me.TermsId)
        Page.AddVariable("%quote_id%", Me.QuoteId)
        Page.AddVariable("%office%", Me.Office)
        ' Billing
        Dim Subtotal As Double = Me.Subtotal
        Dim Tax As Double = Me.SalesTax
        Page.AddVariable("%subtotal%", Format(Subtotal, "$0.00"))
        Page.AddVariable("%tax%", Format(Tax, "$0.00"))
        Page.AddVariable("%total%", Format(Subtotal + Tax, "$0.00"))
        ' Line Items
        Dim Table As New DataTable
        Dim Row As DataRow
        Table.Columns.Add("quantity")
        Table.Columns.Add("part_no")
        Table.Columns.Add("serial_no")
        Table.Columns.Add("description")
        Table.Columns.Add("unit_price")
        Table.Columns.Add("ext_price")
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
        Dim Element As GravityDocument.gElement = Page.GetTableBySource("line_items")
        If Element IsNot Nothing Then
            Element.Table.Data = Table
        End If
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
