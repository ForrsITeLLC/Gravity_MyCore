Imports MyCore.Data

Public Class cPurchaseOrder

    Dim _Id As Integer = 0
    Public No As String = ""
    Public VendorNo As String = ""
    Public ContactName As String = ""
    Public ContactID As Integer = 0
    Public ShipToNo As String = ""
    Public ShipViaId As Integer = 0
    Public Fob As String = ""
    Public RequestedBy As String = ""
    Public Office As Integer = 0
    Public TrackingNo As String = ""
    Public ConfirmationNo As String = ""
    Public DatePlannedShip As Date = Nothing
    Public DateExpected As Date = Nothing
    Public RequisitionNo As String = ""
    Public Resale As Boolean
    Public Notes As String = ""
    Public ShippingCharge As Double = 0
    Public Tax As Double = 0
    Public Terms As String = ""
    Public CreatedBy As String = ""
    Public DateCreated As Date = Now
    Public DateLastUpdated As Date = Nothing
    Public LastUpdatedBy As String = ""
    Public isRMA As Boolean = False
    Public Type As PoType = PoType.Stock
    Public OurOrderNo As Integer
    Public DateOrdered As Date = Nothing
    Public LineItems As POItem()
    Public LineItemsTable As DataTable
    Public UpdatePartsPricing As Boolean = True
    Public Voided As Boolean = False
    Public Released As Boolean = False
    Public AcctRef As String = Nothing
    Public AcctUpload As Boolean = Nothing
    Public ShipViaRef As String = Nothing
    Public ShipToRef As String = Nothing
    Public VendorRef As String = Nothing

    Dim _ShipVia As DataTable
    Dim _Office As DataTable
    Dim _Employees As DataTable

    Dim Database As New MyCore.Data.EasySql("MsSql")

    Public Event Reload()
    Public Event Saved(ByVal PO As cPurchaseOrder)

    Public Enum PoType
        Other = 0
        Stock = 1
        Sales = 2
        Service = 3
        Rental = 4
    End Enum

    Public Structure POItem
        Public Id As Integer
        Public Quantity As Integer
        Public PartNo As String
        Public PartNoRef As String
        Public ItemTypeID As Integer
        Public Description As String
        Public Options As String
        Public UnitPrice As Double
        Public Discount As Double
        Public NetPrice As Double
        Public Received As Integer
    End Structure

    Public ReadOnly Property Id() As Integer
        Get
            Return Me._Id
        End Get
    End Property

    Public ReadOnly Property OfficeRef() As String
        Get
            Dim Ref As String
            Select Case Me.Type
                Case PoType.Sales
                    Ref = Me.Database.GetOne("SELECT sales_ref FROM office WHERE number=" & Me.Database.Escape(Me.Office))
                Case PoType.Service
                    Ref = Me.Database.GetOne("SELECT service_ref FROM office WHERE number=" & Me.Database.Escape(Me.Office))
                Case PoType.Rental
                    Ref = Me.Database.GetOne("SELECT rental_ref FROM office WHERE number=" & Me.Database.Escape(Me.Office))
                Case Else
                    Ref = Me.Database.GetOne("SELECT acct_ref FROM office WHERE number=" & Me.Database.Escape(Me.Office))
            End Select
            If Ref Is DBNull.Value Then
                Ref = Me.Database.GetOne("SELECT acct_ref FROM office WHERE number=" & Me.Database.Escape(Me.Office))
            End If
            Return Ref
        End Get
    End Property

    Public ReadOnly Property Subtotal() As Double
        Get
            Dim Total As Double = 0
            If Me.LineItemsTable IsNot Nothing Then
                For i As Integer = 0 To Me.LineItemsTable.Rows.Count - 1
                    Dim Net As Double = 0
                    Dim Qty As Double = 1
                    If Microsoft.VisualBasic.IsNumeric(Me.LineItemsTable.Rows(i).Item("net_price")) Then
                        Net = Me.LineItemsTable.Rows(i).Item("net_price")
                    End If
                    If Microsoft.VisualBasic.IsNumeric(Me.LineItemsTable.Rows(i).Item("quantity")) Then
                        Qty = Me.LineItemsTable.Rows(i).Item("quantity")
                    End If
                    Total += (Net * Qty)
                Next
            End If
            Return Total
        End Get
    End Property

    Public ReadOnly Property Total() As Double
        Get
            Return Me.Subtotal + Me.ShippingCharge + Me.Tax
        End Get
    End Property

    Public ReadOnly Property Offices() As DataTable
        Get
            Return Me._Office
        End Get
    End Property

    Public ReadOnly Property ShipVia() As DataTable
        Get
            Return Me._ShipVia
        End Get
    End Property

    Public ReadOnly Property Employees() As DataTable
        Get
            Return Me._Employees
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
            Sql &= " WHERE ref_no=" & Me.Database.Escape(Me.No)
            Sql &= " AND department=" & CInt(cInteraction.ReferenceTypes.PO)
            Return Me.Database.GetAll(Sql)
        End Get
    End Property

    Public ReadOnly Property Equipment() As DataTable
        Get
            Dim Sql As String = ""
            Sql &= " SELECT dep_id AS ID"
            Sql &= ", name = (equip.dep_manuf + ' ' + equip.dep_mod + ' (' + equip.dep_ser + ')')"
            Sql &= ", ADDRESS.cst_name AS location_name, location_city = (ADDRESS.cst_city + ', ' + ADDRESS.cst_state)"
            Sql &= " FROM DEPREC equip"
            Sql &= " LEFT OUTER JOIN ADDRESS ON equip.dep_loc = ADDRESS.cst_no "
            Sql &= " WHERE dep_ourpo=" & Me.Database.Escape(Me.No)
            Return Me.Database.GetAll(Sql)
        End Get
    End Property

    Public Sub New(ByRef db As MyCore.Data.EasySql)
        Me.Database = db
        Me.PopulateOffice()
        Me.PopulateShipVia()
        Me.PopuplateEmployees()
    End Sub

    Public Sub New(ByVal Id As Integer)
        Me._Id = Id
        Me.PopulateOffice()
        Me.PopulateShipVia()
        Me.PopuplateEmployees()
    End Sub

    Private Sub PopulateShipVia()
        Me._ShipVia = Me.Database.GetAll("SELECT id, name, sort FROM ship_via ORDER BY sort")
    End Sub

    Private Sub PopulateOffice()
        Me._Office = Me.Database.GetAll("SELECT id, number, name, sort FROM office ORDER BY sort")
    End Sub

    Private Sub PopuplateEmployees()
        Me._Employees = Me.Database.GetAll("SELECT windows_user, (last_name + ', ' + first_name) AS name FROM employee WHERE deactivated=0 ORDER BY last_name, first_name")
    End Sub

    Public Sub Open(ByVal PoNo As String)
        Dim Row As DataRow
        Dim Sql As String = "SELECT po.*, sv.acct_ref AS ship_via_ref, v.acct_ref AS vendor_ref,"
        Sql &= " st.acct_ref AS ship_to_acct_ref"
        Sql &= " FROM purchase_order po"
        Sql &= " LEFT OUTER JOIN ship_via sv on po.ship_via_id=sv.id"
        Sql &= " LEFT OUTER JOIN ADDRESS v ON po.vendor_no=v.cst_no"
        Sql &= " LEFT OUTER JOIN ADDRESS st ON po.ship_to=st.cst_no"
        Sql &= " WHERE po_no=" & Me.Database.Escape(PoNo)
        Row = Me.Database.GetRow(Sql)
        If Me.Database.LastQuery.RowsReturned = 1 Then
            Me._Id = Row.Item("id")
            Me.No = Row.Item("po_no")
            Me.isRMA = Row.Item("rma")
            Me.VendorNo = Row.Item("vendor_no")
            Me.ContactID = Row.Item("contact_id")
            Me.ContactName = Row.Item("contact_name")
            Me.ShipToNo = Row.Item("ship_to")
            Me.DateCreated = Row.Item("po_date")
            Me.ShipViaId = Row.Item("ship_via_id")
            Me.Fob = Row.Item("fob_location")
            Me.RequestedBy = Row.Item("requested_by")
            Me.Office = Row.Item("office")
            Me.TrackingNo = Row.Item("tracking_no")
            Me.ConfirmationNo = Row.Item("confirmation_no")
            Me.DatePlannedShip = IIf(Row.Item("date_planned_ship") Is DBNull.Value, Nothing, Row.Item("date_planned_ship"))
            Me.DateExpected = IIf(Row.Item("date_expected") Is DBNull.Value, Nothing, Row.Item("date_expected"))
            Me.RequisitionNo = Row.Item("requisition_no")
            Me.Resale = Row.Item("for_resale")
            Me.Notes = Row.Item("notes")
            If Me.isRMA Then
                Me.Tax = Row.Item("tax") * -1
                Me.ShippingCharge = Row.Item("shipping_charge") * -1
            Else
                Me.Tax = Row.Item("tax")
                Me.ShippingCharge = Row.Item("shipping_charge")
            End If
            Me.Terms = Row.Item("terms")
            Me.Type = Row.Item("type")
            Me.OurOrderNo = Row.Item("our_order_no")
            Me.DateOrdered = Row.Item("date_ordered")
            Me.UpdatePartsPricing = Row.Item("update_parts_pricing")
            Me.Voided = Row.Item("voided")
            Me.Released = Row.Item("released")
            Me.AcctRef = Me.IfNull(Row.Item("acct_ref"), Nothing)
            Me.AcctUpload = Me.IfNull(Row.Item("acct_upload"), Nothing)
            Me.ShipViaRef = Me.IfNull(Row.Item("ship_via_ref"), Nothing)
            Me.VendorRef = Me.IfNull(Row.Item("vendor_ref"), Nothing)
            Me.ShipToRef = Me.IfNull(Row.Item("ship_to_acct_ref"), Nothing)
        ElseIf Not Me.Database.LastQuery.Successful Then
            Throw New Exception(Me.Database.LastQuery.ErrorMsg)
        Else
            Throw New Exception("Purchase order does not exist.")
        End If
        Try
            Me.OpenLineItems()
            RaiseEvent Reload()
        Catch ex As Exception
            Throw New cException(cException.SeverityRating.Minor, "Couldn't get line items.", ex.ToString)
        End Try
    End Sub

    Public Sub OpenLineItems()
        Dim Sql As String = "SELECT poi.[id], po_no, quantity, poi.part_no, poi.[description], poi.options,"
        If Me.isRMA Then
            Sql &= " (unit_price*-1) AS unit_price, (net_price*-1) AS net_price,"
        Else
            Sql &= " unit_price, net_price,"
        End If
        Sql &= " received, discount,"
        Sql &= " item_master.acct_ref AS part_no_ref,"
        Sql &= " item_master.item_type_id"
        Sql &= " FROM purchase_order_item poi"
        Sql &= " LEFT OUTER JOIN item_master ON item_master.part_no=poi.part_no"
        Sql &= " WHERE po_no=" & Me.Database.Escape(Me.No)
        Dim dt As DataTable = Me.Database.GetAll(Sql)
        If Me.Database.LastQuery.Successful Then
            ReDim Me.LineItems(dt.Rows.Count - 1)
            For i As Integer = 0 To dt.Rows.Count - 1
                Me.LineItems(i).Id = dt.Rows(i).Item("id")
                Me.LineItems(i).Quantity = dt.Rows(i).Item("quantity")
                Me.LineItems(i).PartNo = Me.IsNull(dt.Rows(i).Item("part_no"), "")
                Me.LineItems(i).Description = dt.Rows(i).Item("description")
                Me.LineItems(i).Options = dt.Rows(i).Item("options")
                Me.LineItems(i).UnitPrice = dt.Rows(i).Item("unit_price")
                Me.LineItems(i).Discount = dt.Rows(i).Item("discount")
                If dt.Rows(i).Item("net_price") IsNot DBNull.Value Then
                    Me.LineItems(i).NetPrice = dt.Rows(i).Item("net_price")
                Else
                    Me.LineItems(i).NetPrice = Me.LineItems(i).UnitPrice * (1 - Me.LineItems(i).Discount)
                End If
                Me.LineItems(i).Received = dt.Rows(i).Item("received")
                Me.LineItems(i).PartNoRef = Me.IsNull(dt.Rows(i).Item("part_no_ref"), "")
                Me.LineItems(i).ItemTypeID = Me.IsNull(dt.Rows(i).Item("item_type_id"), 0)
            Next
            Me.LineItemsTable = dt
        Else
            Dim Err As String = Me.Database.LastQuery.ErrorMsg
            System.Windows.Forms.MessageBox.Show(Err)
        End If
    End Sub

    Public Sub Save()
        Dim Sql As String = ""
        If Me._Id = 0 Then
            Sql &= "INSERT INTO purchase_order (po_no, vendor_no, contact_name, contact_id, ship_to, po_date,"
            Sql &= " ship_via_id, fob_location, requested_by, office, tracking_no, confirmation_no, date_planned_ship, "
            Sql &= " date_expected, requisition_no, for_resale, notes, tax, shipping_charge, terms, type,"
            Sql &= " our_order_no, date_ordered, update_parts_pricing, voided, rma, released)"
            Sql &= " VALUES ("
            Sql &= " @po_no, @vendor_no, @contact_name, @contact_id, @ship_to, @po_date, @ship_via_id,"
            Sql &= " @fob_location, @requested_by, @office, @tracking_no, @confirmation_no,"
            Sql &= " @date_planned_ship, @date_expected, @requisition_no, @for_resale,"
            Sql &= " @notes, @tax, @shipping_charge, @terms, @type, @our_order_no,"
            Sql &= " @date_ordered, @update_parts_pricing, @voided, @rma, @released)"
        Else
            Sql &= "UPDATE purchase_order "
            Sql &= " SET po_no=@po_no, vendor_no=@vendor_no, contact_name=@contact_name, contact_id=@contact_id, "
            Sql &= " ship_to=@ship_to, po_date=@po_date, ship_via_id=@ship_via_id, fob_location=@fob_location, "
            Sql &= " requested_by=@requested_by, office=@office, tracking_no=@tracking_no, "
            Sql &= " confirmation_no=@confirmation_no, date_planned_ship=@date_planned_ship, "
            Sql &= " date_expected=@date_expected, requisition_no=@requisition_no, for_resale=@for_resale, "
            Sql &= " notes=@notes, tax=@tax, shipping_charge=@shipping_charge, terms=@terms, type=@type,"
            Sql &= " our_order_no=@our_order_no, date_ordered=@date_ordered, update_parts_pricing=@update_parts_pricing,"
            Sql &= " voided=@voided, rma=@rma, released=@released"
            Sql &= " WHERE id=@id"
        End If
        Sql = Sql.Replace("@po_no", Me.Database.Escape(Me.No))
        Sql = Sql.Replace("@vendor_no", Me.Database.Escape(Me.VendorNo))
        Sql = Sql.Replace("@contact_id", Me.Database.Escape(Me.ContactID))
        Sql = Sql.Replace("@contact_name", Me.Database.Escape(Me.ContactName))
        Sql = Sql.Replace("@ship_to", Me.Database.Escape(Me.ShipToNo))
        Sql = Sql.Replace("@po_date", Me.Database.Escape(Me.DateCreated))
        Sql = Sql.Replace("@ship_via_id", Me.Database.Escape(Me.ShipViaId))
        Sql = Sql.Replace("@fob_location", Me.Database.Escape(Me.Fob))
        Sql = Sql.Replace("@requested_by", Me.Database.Escape(Me.RequestedBy))
        Sql = Sql.Replace("@office", Me.Database.Escape(Me.Office))
        Sql = Sql.Replace("@tracking_no", Me.Database.Escape(Me.TrackingNo))
        Sql = Sql.Replace("@confirmation_no", Me.Database.Escape(Me.ConfirmationNo))
        If Me.isRMA Then
            If Me.DatePlannedShip = Nothing Then
                Sql = Sql.Replace("@date_planned_ship", "NULL")
            Else
                Sql = Sql.Replace("@date_planned_ship", Me.Database.Escape(Me.DatePlannedShip))
            End If
            Sql = Sql.Replace("@date_expected", "NULL")
            Sql = Sql.Replace("@tax", Me.Database.Escape(Me.Tax * -1))
            Sql = Sql.Replace("@shipping_charge", Me.Database.Escape(Me.ShippingCharge * -1))
        Else
            Sql = Sql.Replace("@date_planned_ship", Me.Database.Escape(Me.DatePlannedShip))
            Sql = Sql.Replace("@date_expected", Me.Database.Escape(Me.DateExpected))
            Sql = Sql.Replace("@tax", Me.Database.Escape(Me.Tax))
            Sql = Sql.Replace("@shipping_charge", Me.Database.Escape(Me.ShippingCharge))
        End If

        Sql = Sql.Replace("@requisition_no", Me.Database.Escape(Me.RequisitionNo))
        Sql = Sql.Replace("@for_resale", Me.Database.Escape(Me.Resale))
        Sql = Sql.Replace("@notes", Me.Database.Escape(Me.Notes))
        Sql = Sql.Replace("@terms", Me.Database.Escape(Me.Terms))
        Sql = Sql.Replace("@type", Me.Database.Escape(CInt(Me.Type)))
        Sql = Sql.Replace("@our_order_no", Me.Database.Escape(Me.OurOrderNo))
        Sql = Sql.Replace("@date_ordered", Me.Database.Escape(Me.DateOrdered))
        Sql = Sql.Replace("@update_parts_pricing", Me.Database.Escape(Me.UpdatePartsPricing))
        Sql = Sql.Replace("@voided", Me.Database.Escape(Me.Voided))
        Sql = Sql.Replace("@rma", Me.Database.Escape(Me.isRMA))
        Sql = Sql.Replace("@released", Me.Database.Escape(Me.Released))
        If Me._Id > 0 Then
            Sql = Sql.Replace("@id", Me._Id)
            Me.Database.Execute(Sql)
        Else
            Me.Database.Execute(Sql)
        End If
        If Me.Database.LastQuery.Successful Then
            RaiseEvent Saved(Me)
            If Me._Id = 0 Then
                Me.IncrementNextNumber()
                Me.Open(Me.No)
            Else
                Me.SaveLineItems()
                Me.Open(Me.No)
            End If
        Else
            Throw New Exception(Me.Database.LastQuery.ErrorMsg & " " & Me.Database.LastQuery.CommandText)
        End If
    End Sub

    Private Sub SaveLineItems()
        Dim dr As DataRow
        For Each dr In Me.LineItemsTable.Rows
            Dim Params As New Collection
            Dim PartNo As String = Me.IsNull(dr.Item("part_no"), "")
            If dr.RowState = DataRowState.Added Then
                If dr.Item("quantity") > 0 Then
                    If dr.Item("description").ToString.Length > 0 Or dr.Item("part_no").ToString.Length > 0 Then
                        Dim Sql As String = "INSERT INTO purchase_order_item (po_no, quantity, part_no, [description], options, unit_price, received, discount, net_price)"
                        Sql &= " VALUES (@po_no, @quantity, @part_no, @description, @options, "
                        Sql &= Me.Database.ToCurrency("@unit_price") & ", @received, @discount, " & Me.Database.ToCurrency("@net_price") & ")"
                        Sql = Sql.Replace("@po_no", Me.Database.Escape(Me.No))
                        Sql = Sql.Replace("@quantity", Me.IsNull(dr.Item("quantity"), 0))
                        Sql = Sql.Replace("@part_no", Me.Database.Escape(PartNo))
                        Sql = Sql.Replace("@description", Me.Database.Escape(Me.IsNull(dr.Item("description"), "")))
                        Sql = Sql.Replace("@options", Me.Database.Escape(Me.IsNull(dr.Item("options"), "")))
                        If Me.isRMA Then
                            Sql = Sql.Replace("@unit_price", Me.Database.Escape(IIf(dr.Item("unit_price") Is DBNull.Value, 0, dr.Item("unit_price")) * -1))
                            Sql = Sql.Replace("@net_price", Me.Database.Escape(Me.IsNull(dr.Item("net_price"), 0) * -1))
                        Else
                            Sql = Sql.Replace("@unit_price", Me.Database.Escape(IIf(dr.Item("unit_price") Is DBNull.Value, 0, dr.Item("unit_price"))))
                            Sql = Sql.Replace("@net_price", Me.Database.Escape(Me.IsNull(dr.Item("net_price"), 0)))
                        End If
                        Sql = Sql.Replace("@discount", Me.Database.Escape(Me.IsNull(dr.Item("discount"), 0)))
                        Sql = Sql.Replace("@received", Me.Database.Escape(IIf(dr.Item("received") Is DBNull.Value, 0, dr.Item("received"))))
                        Me.Database.Execute(Sql)

                        If PartNo.Length > 0 And Me.UpdatePartsPricing Then
                            Me.UpdateMasterPartRecord(dr)
                        End If
                        If Not Me.Database.LastQuery.Successful Then
                            Dim Msg As String = Me.Database.LastQuery.ErrorMsg
                        End If
                    End If
                End If
            ElseIf dr.RowState = DataRowState.Modified Then
                If dr.Item("quantity") = 0 Then
                    Me.Database.Execute("DELETE FROM purchase_order_item WHERE id=" & dr.Item("id"))
                Else
                    Dim Sql As String = "UPDATE  purchase_order_item "
                    Sql &= " SET quantity=@quantity, part_no=@part_no, [description]=@description, "
                    Sql &= " options=@options, unit_price=" & Me.Database.ToCurrency("@unit_price") & ", "
                    Sql &= " received=@received, discount=@discount, "
                    Sql &= " net_price=" & Me.Database.ToCurrency("@net_price")
                    Sql &= " WHERE [id]=@id"
                    Sql = Sql.Replace("@id", dr.Item("id"))
                    Sql = Sql.Replace("@quantity", Me.IsNull(dr.Item("quantity"), 0))
                    Sql = Sql.Replace("@part_no", Me.Database.Escape(PartNo))
                    Sql = Sql.Replace("@description", Me.Database.Escape(Me.IsNull(dr.Item("description"), "")))
                    Sql = Sql.Replace("@options", Me.Database.Escape(Me.IsNull(dr.Item("options"), "")))
                    If Me.isRMA Then
                        Sql = Sql.Replace("@unit_price", Me.Database.Escape(IIf(dr.Item("unit_price") Is DBNull.Value, 0, dr.Item("unit_price")) * -1))
                        Sql = Sql.Replace("@net_price", Me.Database.Escape(Me.IsNull(dr.Item("net_price"), 0) * -1))
                    Else
                        Sql = Sql.Replace("@unit_price", Me.Database.Escape(IIf(dr.Item("unit_price") Is DBNull.Value, 0, dr.Item("unit_price"))))
                        Sql = Sql.Replace("@net_price", Me.Database.Escape(Me.IsNull(dr.Item("net_price"), 0)))
                    End If
                    Sql = Sql.Replace("@discount", Me.Database.Escape(Me.IsNull(dr.Item("discount"), 0)))
                    Sql = Sql.Replace("@received", Me.Database.Escape(IIf(dr.Item("received") Is DBNull.Value, 0, dr.Item("received"))))
                    Me.Database.Execute(Sql)
                    If PartNo.Length > 0 And Me.UpdatePartsPricing Then
                        Me.UpdateMasterPartRecord(dr)
                    End If
                End If
            End If
            If Not Me.Database.LastQuery.Successful Then
                MsgBox(Me.Database.LastQuery.ErrorMsg)
            End If
        Next
    End Sub

    Private Sub UpdateMasterPartRecord(ByVal Row As DataRow)
        If Not Me.isRMA Then
            Dim Price As Double = Me.IsNull(Row.Item("unit_price"), 0)
            Dim Cost As Double = Me.IsNull(Row.Item("net_price"), 0)
            If Price > 0 Or Cost > 0 Then
                Dim Sql As String = ""
                Sql = "UPDATE item_master SET "
                If Cost > 0 Then
                    Sql &= " cost=" & Me.Database.ToCurrency(Cost)
                End If
                If Price > 0 Then
                    If Cost > 0 Then
                        Sql &= ", "
                    End If
                    Sql &= " list_price=" & Me.Database.ToCurrency(Price)
                End If
                Sql &= " WHERE part_no=" & Me.Database.Escape(Row.Item("part_no"))
                Me.Database.Execute(Sql)
            End If
        End If
    End Sub

    Private Function IsNull(ByVal Value As Object, ByVal ReturnVal As String) As String
        If Value Is DBNull.Value Then
            Return ReturnVal
        Else
            Return Value
        End If
    End Function

    Public Function ToGravityDocument(ByVal Template As String) As GravityDocument.gDocument
        ' If no template specified
        If Template.Length = 0 Then
            Dim id As Integer
            If isRMA Then
                id = Me.Database.GetOne("SELECT value FROM settings WHERE property='Template RMA'")
            Else
                id = Me.Database.GetOne("SELECT value FROM settings WHERE property='Template PO'")
            End If
            Template = Me.Database.GetOne("SELECT html FROM template WHERE id=" & id)
        End If
        ' Create Gravity Document
        Dim Doc As New GravityDocument.gDocument(Me.Database.GetOne("SELECT value FROM settings WHERE property='Page Height in Pixels'"))
        Doc.LoadXml(Template)
        ' Settings
        Doc.FormType = GravityDocument.gDocument.FormTypes.PurchaseOrder
        Doc.ReferenceID = Me.No
        ' Put in variables
        ' BILL TO
        Dim Sql As String = "SELECT cst_name, cst_addr1, cst_addr2, cst_city, cst_state, cst_zip,"
        Sql &= " cst_phone, cst_fax"
        Sql &= " FROM ADDRESS"
        Sql &= " WHERE cst_no = " & Me.Database.Escape(Me.VendorNo)
        Dim Vendor As DataRow = Me.Database.GetRow(Sql)
        If Not Me.Database.LastQuery.Successful Then
            Dim Err As String = Me.Database.LastQuery.ErrorMsg
        End If
        ' Create address strings
        Dim VendorAddress As String = ""
        ' If bill to has a billing address use it
        VendorAddress = IfNull(Vendor.Item("cst_addr1"))
        If Me.IfNull(Vendor.Item("cst_addr2")).Length > 0 Then
            VendorAddress &= ControlChars.CrLf & IfNull(Vendor.Item("cst_addr2"))
        End If
        ' SHIP TO
        Dim ShipTo As DataRow
        Dim ShipToAddress As String = ""
        If Me.isRMA Then
            ShipTo = Vendor
            ShipToAddress = VendorAddress
        Else
            Sql = "SELECT cst_name, cst_addr1, cst_addr2, cst_city, cst_state, cst_zip, cst_phone, cst_fax"
            Sql &= " FROM ADDRESS WHERE cst_no=" & Me.Database.Escape(Me.ShipToNo)
            ShipTo = Me.Database.GetRow(Sql)
            ShipToAddress = IfNull(ShipTo.Item("cst_addr1"))
            If Me.IfNull(ShipTo.Item("cst_addr2")).Length > 0 Then
                ShipToAddress &= ControlChars.CrLf & IfNull(ShipTo.Item("cst_addr2"))
            End If
        End If
        ' Replace variables
        Dim Page As GravityDocument.gPage = Doc.GetPage(1)
        Page.AddVariable("%po_no%", Me.No)
        Page.AddVariable("%order_date%", Format(Me.DateOrdered, "MM/dd/yy"))
        Page.AddVariable("%ship_date%", Format(Me.DatePlannedShip, "MM/dd/yy"))
        Page.AddVariable("%expected_date%", Format(Me.DateExpected, "MM/dd/yy"))
        Page.AddVariable("%ship_to_name%", ShipTo.Item("cst_name"))
        Page.AddVariable("%ship_to_address%", ShipToAddress)
        Page.AddVariable("%ship_to_city%", ShipTo.Item("cst_city"))
        Page.AddVariable("%ship_to_state%", ShipTo.Item("cst_state"))
        Page.AddVariable("%ship_to_zip%", Me.IsNull(ShipTo.Item("cst_zip"), ""))
        Page.AddVariable("%ship_to_phone%", Me.IsNull(ShipTo.Item("cst_phone"), ""))
        Page.AddVariable("%ship_to_fax%", Me.IsNull(ShipTo.Item("cst_fax"), ""))
        Page.AddVariable("%vendor_name%", Vendor.Item("cst_name"))
        Page.AddVariable("%vendor_address%", VendorAddress)
        Page.AddVariable("%vendor_city%", Vendor.Item("cst_city"))
        Page.AddVariable("%vendor_state%", Vendor.Item("cst_state"))
        Page.AddVariable("%vendor_zip%", Me.IsNull(Vendor.Item("cst_zip"), ""))
        Page.AddVariable("%vendor_phone%", Me.IsNull(Vendor.Item("cst_phone"), ""))
        Page.AddVariable("%vendor_fax%", Me.IsNull(Vendor.Item("cst_fax"), ""))
        Page.AddVariable("%vendor_no%", Me.VendorNo)
        Page.AddVariable("%ship_to_no%", Me.ShipToNo)
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
        ' Order details
        Page.AddVariable("%fob%", Me.Fob)
        Page.AddVariable("%ship_via%", Me.Database.GetOne("SELECT name FROM ship_via WHERE id=" & Me.ShipViaId))
        Page.AddVariable("%requisitor%", Me.RequestedBy)
        Page.AddVariable("%requisition_no%", Me.RequisitionNo)
        Page.AddVariable("%terms%", Me.Terms)
        Page.AddVariable("%notes%", Me.Notes)
        Page.AddVariable("%confirmation_no%", Me.ConfirmationNo)
        Page.AddVariable("%contact%", Me.ContactName)
        Page.AddVariable("%office%", Me.Office)
        Page.AddVariable("%our_order_no%", Me.OurOrderNo)
        Page.AddVariable("%freight%", Me.ShippingCharge)
        Page.AddVariable("%tax_p%", Me.Tax)
        Page.AddVariable("%tracking_no%", Me.TrackingNo)
        Page.AddVariable("%resale_yn%", IIf(Me.Resale, "Y", "N"))
        Page.AddVariable("%resale_yesno%", IIf(Me.Resale, "Yes", "No"))
        Page.AddVariable("%resale%", IIf(Me.Resale, "For Resale", "Not For Resale"))
        ' Billing
        Page.AddVariable("%subtotal%", Format(Me.Subtotal, "$0.00"))
        Page.AddVariable("%total%", Format(Me.Total, "$0.00"))
        ' Line Items
        Dim Table As New DataTable
        Table.Columns.Add("quantity")
        Table.Columns.Add("part_no")
        Table.Columns.Add("description")
        Table.Columns.Add("options")
        Table.Columns.Add("list_price")
        Table.Columns.Add("discount")
        Table.Columns.Add("net_price")
        Table.Columns.Add("ext_price")
        For Each Row As DataRow In Me.LineItemsTable.Rows
            Dim r As DataRow = Table.NewRow
            r.Item("quantity") = Row.Item("quantity")
            r.Item("part_no") = Row.Item("part_no")
            r.Item("options") = Row.Item("options")
            r.Item("description") = Row.Item("description")
            r.Item("list_price") = Row.Item("unit_price")
            r.Item("discount") = Row.Item("discount")
            r.Item("net_price") = Row.Item("net_price")
            r.Item("ext_price") = Row.Item("net_price") * Row.Item("quantity")
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

    Public Sub IncrementNextNumber()
        Me.Database.Execute("UPDATE next_number SET number=number+1 WHERE name='po'")
    End Sub

    Public Function GetNextNumber() As Integer
        Return Me.Database.GetOne("SELECT number FROM next_number WHERE name='po'")
    End Function


End Class

