Public Class cItemMaster

    Public Manufacturer As String = ""
    Public Model As String = ""
    Public PartNo As String = ""
    Public CategoryId As Integer = 0
    Public OrderDescription As String = ""
    Public Notes As String = ""
    Public VendorNo As String = ""
    Public LaborWarrantyDays As Integer = 0
    Public PartsWarrantyDays As Integer = 0
    Public ListPrice As Double = 0.0
    Public Cost As Double = 0.0
    Public Freight As Double = 0.0
    Public NoReorder As Boolean = False
    Public AItem As Boolean = False
    Public Prime As Boolean = False
    Public Inactive As Boolean = False
    Public GroupId As Integer = 0
    Public TaxableValue As Double = 0
    Public ItemType As Integer = 1
    Public InternalOnly As Boolean = False
    Public CostAccount As String = ""
    Public SalesAccount As String = ""
    Public AssetAccount As String = ""
    Public TaxStatus As Integer = 1
    Public VendorPartNo As String = ""

    Public LastUpdatedBy As String = ""
    Public CreatedBy As String = ""

    Public AcctRef As String = Nothing
    Public AcctUpload As Boolean = Nothing
    Public VendorRef As String = Nothing
    Public TaxStatusRef As String = Nothing
    Public SalesAccountRef As String = Nothing
    Public CostAccountRef As String = Nothing
    Public AssetAccountRef As String = Nothing

    Dim _Id As Integer = 0
    Dim _Categories As DataTable
    Dim _Stations As DataTable
    Dim _Groups As DataTable
    Dim Database As New MyCore.Data.EasySql("MsSql")

    Public Event Reload()
    Public Event NewItemMasterCreated()

    Public Enum InventoryChangeReason
        Sold = 1
        Purchased = 2
        Trade = 3
        Correction = 4
        Retired = 5
        Moved = 6
        RentalDelivered = 7
        RentalReturned = 8
        Returned = 9
    End Enum

    Public Enum ItemTypesEnum
        Inventory = 1
        Labor = 2
        Travel = 3
        NonInventory = 4
        OtherCharge = 5
        Discount = 6
        TaxItem = 7
        Special = 8
    End Enum

    Public ReadOnly Property Id() As Integer
        Get
            Return Me._Id
        End Get
    End Property

    Public ReadOnly Property Categories() As DataTable
        Get
            Return Me._Categories
        End Get
    End Property

    Public ReadOnly Property Stations() As DataTable
        Get
            Return Me._Stations
        End Get
    End Property

    Public ReadOnly Property Taxable() As Boolean
        Get
            Return Me.Database.GetOne("SELECT taxable FROM tax_status WHERE id=" & Me.TaxStatus)
        End Get
    End Property

    Public ReadOnly Property Groups() As DataTable
        Get
            Return Me._Groups
        End Get
    End Property

    Public ReadOnly Property SerialNumbersInStock() As DataTable
        Get
            If Me.Prime Then
                Dim Sql As String = ""
                Sql &= "SELECT dep_id AS id, dep_ser AS serial_no, dep_loc AS location,"
                Sql &= " l.cst_city + ', ' + l.cst_state AS location_city,"
                Sql &= " s.name AS station, station_id, dep_stat AS status,"
                Sql &= " (SELECT TOP 1 dstat_chngt FROM DEPRECSTATUS"
                Sql &= " WHERE dstat_dep_id=dep_id ORDER BY dstat_chngt DESC"
                Sql &= " ) AS date_last_updated"
                Sql &= " FROM DEPREC"
                Sql &= " LEFT OUTER JOIN ADDRESS l ON dep_loc=cst_no"
                Sql &= " INNER JOIN station s ON s.id=station_id"
                Sql &= " WHERE part_no=" & Me.Database.Escape(Me.PartNo)
                Sql &= " AND dep_stat IN (SELECT code FROM equipment_status WHERE in_stock=1)"
                Sql &= " ORDER BY dep_ser"
                Return Me.Database.GetAll(Sql)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property SerialNumbersInStock(ByVal StationID As Integer) As DataTable
        Get
            If Me.Prime Then
                Dim Sql As String = ""
                Sql &= "SELECT dep_id AS id, dep_ser AS serial_no, dep_loc AS location,"
                Sql &= " l.cst_city + ', ' + l.cst_state AS location_city,"
                Sql &= " s.name AS station, station_id, dep_stat AS status,"
                Sql &= " (SELECT TOP 1 dstat_chngt FROM DEPRECSTATUS"
                Sql &= " WHERE dstat_dep_id=dep_id ORDER BY dstat_chngt DESC"
                Sql &= " ) AS date_last_updated"
                Sql &= " FROM DEPREC"
                Sql &= " LEFT OUTER JOIN ADDRESS l ON dep_loc=cst_no"
                Sql &= " INNER JOIN station s ON s.id=station_id"
                Sql &= " WHERE part_no=" & Me.Database.Escape(Me.PartNo)
                Sql &= " AND station_id=" & StationID
                Sql &= " AND dep_stat IN (SELECT code FROM equipment_status WHERE in_stock=1)"
                Sql &= " ORDER BY dep_ser"
                Return Me.Database.GetAll(Sql)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property TotalInStock() As Integer
        Get
            If Me.ItemType = 1 Then
                Dim Sql As String = ""
                If Me.Prime Then
                    Sql &= "SELECT COUNT(dep_id) AS total_stock"
                    Sql &= " FROM DEPREC"
                    Sql &= " WHERE part_no=" & Me.Database.Escape(Me.PartNo)
                    Sql &= " AND dep_stat IN (SELECT code FROM equipment_status WHERE in_stock=1)"
                    Return Me.Database.GetOne(Sql)
                Else
                    Sql = "SELECT SUM(quantity) AS total_stock FROM item_inventory WHERE part_no=" & Me.Database.Escape(Me.PartNo)
                    Return Me.Database.GetOne(Sql)
                End If
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property TotalInStockByStockRoom() As DataTable
        Get
            If Me.ItemType = 1 Then
                Dim Sql As String = ""
                If Me.Prime Then
                    Sql &= "SELECT COUNT(dep_id) AS quantity, station_id, station.name AS stock_room"
                    Sql &= " FROM DEPREC"
                    Sql &= " INNER JOIN station ON station_id=station.id"
                    Sql &= " WHERE part_no=" & Me.Database.Escape(Me.PartNo)
                    Sql &= " AND dep_stat IN (SELECT code FROM equipment_status WHERE in_stock=1)"
                    Sql &= " GROUP BY station_id, station.name"
                    Return Me.Database.GetAll(Sql)
                Else
                    Sql = "SELECT SUM(quantity) AS quantity, station_id, s.name AS stock_room,"
                    Sql &= " (SELECT TOP 1 date_created FROM item_inventory_adjustment"
                    Sql &= " WHERE part_no=" & Me.Database.Escape(Me.PartNo)
                    Sql &= " AND (to_station=station_id OR from_station=station_id)"
                    Sql &= " ORDER BY date_created DESC"
                    Sql &= " ) AS date_last_updated"
                    Sql &= " FROM item_inventory"
                    Sql &= " INNER JOIN station s ON station_id=s.id"
                    Sql &= " WHERE part_no = " & Me.Database.Escape(Me.PartNo)
                    Sql &= " GROUP BY station_id, s.name"
                    Dim Table As DataTable = Me.Database.GetAll(Sql)
                    If Me.Database.LastQuery.Successful Then
                        Return Table
                    Else
                        Dim Err As String = Me.Database.LastQuery.ErrorMsg
                        Return Nothing
                    End If
                End If
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property RentalsReserved(ByVal RentStart As Date, ByVal RentEnd As Date, ByVal OrderId As Integer) As DataTable
        Get
            Dim Sql As String = ""
            Sql &= "SELECT SUM(quantity) AS qty, station_id, station.name AS stock_room"
            Sql &= " FROM rental_order_item"
            Sql &= " INNER JOIN station ON station_id=station.id"
            Sql &= " WHERE rental_order_id IN "
            Sql &= "(SELECT id FROM rental_order WHERE date_start < " & Me.Database.Escape(RentEnd)
            Sql &= " AND DATEADD(day, duration_days, date_start) > " & Me.Database.Escape(RentStart)
            Sql &= " AND rental_order_id <> " & OrderId & ")"
            Sql &= " AND part_no=" & Me.Database.Escape(Me.PartNo)
            Dim Out As DataTable = Me.Database.GetAll(Sql)
            Return Out
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            If Me.OrderDescription.Length > 0 Then
                Return Me.OrderDescription
            Else
                Return Me.Manufacturer & " " & Me.Model
            End If
        End Get
    End Property

    Public Sub New(ByRef db As MyCore.Data.EasySql)
        Me.Database = db
        Me.PopulateCategories()
        Me.PopulateStations()
        Me.PopulateGroups()
    End Sub

#Region "Populate Choices"

    Private Sub PopulateCategories()
        Me._Categories = Me.Database.GetAll("SELECT id, name, sort FROM item_category ORDER BY sort")
    End Sub

    Private Sub PopulateStations()
        Me._Stations = Me.Database.GetAll("SELECT id, name, sort FROM station ORDER BY sort")
    End Sub

    Private Sub PopulateGroups()
        Me._Groups = Me.Database.GetAll("SELECT id, name, sort FROM item_group ORDER BY sort")
    End Sub

#End Region


    Public Sub Open(ByVal PartNo As String)
        Dim Row As DataRow
        Dim Sql As String = ""
        Sql = "SELECT im.*, ts.acct_ref AS tax_status_ref, v.acct_ref AS vendor_ref,"
        Sql &= " igl.acct_ref AS igl_ref, cgl.acct_ref AS cgl_ref, agl.acct_ref AS agl_ref"
        Sql &= " FROM item_master im"
        Sql &= " LEFT OUTER JOIN tax_status ts on im.tax_status_id=ts.id"
        Sql &= " LEFT OUTER JOIN ADDRESS v ON im.vendor_no=v.cst_no"
        Sql &= " LEFT OUTER JOIN gl_account igl ON im.income_gl=igl.account_no"
        Sql &= " LEFT OUTER JOIN gl_account cgl ON im.cost_gl=cgl.account_no"
        Sql &= " LEFT OUTER JOIN gl_account agl ON im.asset_gl=agl.account_no"
        Sql &= " WHERE part_no = " & Me.Database.Escape(PartNo)
        Row = Me.Database.GetRow(Sql)
        If Me.Database.LastQuery.RowsReturned = 1 Then
            Me._Id = Row.Item("id")
            Me.PartNo = Row.Item("part_no")
            Me.Manufacturer = Me.IfNull(Row.Item("manufacturer"))
            Me.Model = Me.IfNull(Row.Item("model"))
            Me.Notes = Me.IfNull(Row.Item("description"))
            Me.OrderDescription = Me.IfNull(Row.Item("order_description"))
            Me.VendorNo = Me.IfNull(Row.Item("vendor_no"))
            Me.PartsWarrantyDays = Me.IfNull(Row.Item("parts_warranty_days"), 0)
            Me.LaborWarrantyDays = Me.IfNull(Row.Item("labor_warranty_days"), 0)
            Me.ItemType = Row.Item("item_type_id")
            Me.CategoryId = Row.Item("item_category_id")
            Me.GroupId = Row.Item("item_group_id")
            Me.ListPrice = Row.Item("list_price")
            Me.Cost = Row.Item("cost")
            Me.Freight = Row.Item("freight")
            Me.Prime = Row.Item("prime")
            Me.NoReorder = Row.Item("no_reorder")
            Me.AItem = Row.Item("a_item")
            Me.Inactive = Row.Item("inactive")
            Me.ItemType = Row.Item("item_type_id")
            Me.InternalOnly = Row.Item("internal_only")
            Me.CostAccount = Row.Item("cost_gl")
            Me.SalesAccount = Row.Item("income_gl")
            Me.AssetAccount = Row.Item("asset_gl")
            Me.TaxStatus = Row.Item("tax_status_id")
            Me.VendorPartNo = Row.Item("vendor_part_no")
            Me.LastUpdatedBy = Row.Item("last_updated_by")
            Me.CreatedBy = Row.Item("created_by")
            Me.AcctRef = Me.IfNull(Row.Item("acct_ref"), Nothing)
            Me.AcctUpload = Me.IfNull(Row.Item("acct_upload"), Nothing)
            Me.VendorRef = Me.IfNull(Row.Item("vendor_ref"), Nothing)
            Me.TaxStatusRef = Me.IfNull(Row.Item("tax_status_ref"), Nothing)
            Me.SalesAccountRef = Me.IfNull(Row.Item("igl_ref"), Nothing)
            Me.CostAccountRef = Me.IfNull(Row.Item("cgl_ref"), Nothing)
            Me.AssetAccountRef = Me.IfNull(Row.Item("agl_ref"), Nothing)
            RaiseEvent Reload()
        Else
            ' Look for X-Ref
            Dim Ref As String = Me.Database.GetOne("SELECT new_part_no FROM part_no_xref WHERE old_part_no=" & Me.Database.Escape(PartNo))
            If Me.Database.LastQuery.RowsReturned = 1 And Not PartNo = Ref Then
                Me.Open(Ref)
            Else
                Throw New Exception("Part number not found.")
            End If
        End If
    End Sub

    Private Function IfNull(ByVal Value As Object, Optional ByVal DefaultVal As Object = "") As Object
        If Value Is DBNull.Value Then
            Return DefaultVal
        Else
            Return Value
        End If
    End Function


    Public Sub Save()
        Dim Params As New Collection
        Dim Sql As String = ""
        If Me._Id > 0 Then
            Sql &= "UPDATE item_master SET"
            Sql &= " manufacturer=@manufacturer,"
            Sql &= " model=@model, "
            Sql &= " description=@notes,"
            Sql &= " vendor_no=@vendor_no,"
            Sql &= " parts_warranty_days=@parts_warranty_days,"
            Sql &= " labor_warranty_days=@labor_warranty_days,"
            Sql &= " item_category_id=@category_id, "
            Sql &= " list_price=@list_price, "
            Sql &= " prime=@prime, "
            Sql &= " no_reorder=@no_reorder,"
            Sql &= " a_item=@a_item, "
            Sql &= " inactive=@inactive, "
            Sql &= " cost=@our_cost, "
            Sql &= " freight=@freight, "
            Sql &= " item_group_id=@group_id, "
            Sql &= " order_description=@order_description,"
            Sql &= " item_type_id=@item_type_id, "
            Sql &= " internal_only=@internal_only, "
            Sql &= " cost_gl=@cost_gl, "
            Sql &= " income_gl=@income_gl, "
            Sql &= " asset_gl=@asset_gl, "
            Sql &= " tax_status_id=@tax_status_id, "
            Sql &= " vendor_part_no=@vendor_part_no, "
            Sql &= " date_last_updated=" & Me.Database.Escape(Now) & ", "
            Sql &= " last_updated_by=" & Me.Database.Escape(Me.LastUpdatedBy) & ", "
            Sql &= " acct_ref=" & Me.Database.Escape(Me.AcctRef)
            Sql &= " WHERE part_no=@part_no"
        Else
            Sql &= "INSERT INTO item_master ("
            Sql &= " part_no, manufacturer, model, [description], vendor_no, parts_warranty_days, labor_warranty_days,"
            Sql &= " item_category_id, list_price, prime, no_reorder, a_item, inactive, cost, freight, item_group_id,"
            Sql &= " order_description, item_type_id, internal_only, cost_gl, income_gl, asset_gl, tax_status_id, vendor_part_no,"
            Sql &= " date_last_updated, date_created, last_updated_by, created_by"
            Sql &= " ) VALUES ("
            Sql &= " @part_no, @manufacturer, @model, @notes, @vendor_no,"
            Sql &= " @parts_warranty_days, @labor_warranty_days,"
            Sql &= " @category_id, @list_price, @prime, @no_reorder,"
            Sql &= " @a_item, @inactive, @our_cost, @freight, "
            Sql &= " @group_id, @order_description, @item_type_id, @internal_only,"
            Sql &= " @cost_gl, @income_gl, @asset_gl, @tax_status_id, @vendor_part_no, "
            Sql &= Me.Database.Escape(Now) & ", "
            Sql &= Me.Database.Escape(Now) & ", "
            Sql &= Me.Database.Escape(Me.LastUpdatedBy) & ", "
            Sql &= Me.Database.Escape(Me.LastUpdatedBy)
            Sql &= " )"
        End If
        Sql = Sql.Replace("@part_no", Me.Database.Escape(Me.PartNo))
        Sql = Sql.Replace("@manufacturer", Me.Database.Escape(Me.Manufacturer))
        Sql = Sql.Replace("@model", Me.Database.Escape(Me.Model))
        Sql = Sql.Replace("@notes", Me.Database.Escape(Me.Notes))
        Sql = Sql.Replace("@vendor_no", Me.Database.Escape(Me.VendorNo))
        Sql = Sql.Replace("@parts_warranty_days", Me.Database.Escape(Me.PartsWarrantyDays))
        Sql = Sql.Replace("@labor_warranty_days", Me.Database.Escape(Me.LaborWarrantyDays))
        Sql = Sql.Replace("@category_id", Me.Database.Escape(Me.CategoryId))
        Sql = Sql.Replace("@list_price", Me.Database.Escape(Me.ListPrice))
        Sql = Sql.Replace("@prime", Me.Database.Escape(Me.Prime))
        Sql = Sql.Replace("@no_reorder", Me.Database.Escape(Me.NoReorder))
        Sql = Sql.Replace("@a_item", Me.Database.Escape(Me.AItem))
        Sql = Sql.Replace("@inactive", Me.Database.Escape(Me.Inactive))
        Sql = Sql.Replace("@our_cost", Me.Database.Escape(Me.Cost))
        Sql = Sql.Replace("@freight", Me.Database.Escape(Me.Freight))
        Sql = Sql.Replace("@group_id", Me.Database.Escape(Me.GroupId))
        Sql = Sql.Replace("@order_description", Me.Database.Escape(Me.OrderDescription))
        Sql = Sql.Replace("@item_type_id", Me.Database.Escape(Me.ItemType))
        Sql = Sql.Replace("@internal_only", Me.Database.Escape(Me.InternalOnly))
        Sql = Sql.Replace("@cost_gl", Me.Database.Escape(Me.CostAccount))
        Sql = Sql.Replace("@income_gl", Me.Database.Escape(Me.SalesAccount))
        Sql = Sql.Replace("@asset_gl", Me.Database.Escape(Me.AssetAccount))
        Sql = Sql.Replace("@tax_status_id", Me.Database.Escape(Me.TaxStatus))
        Sql = Sql.Replace("@vendor_part_no", Me.Database.Escape(Me.VendorPartNo))
        Me.Database.Execute(Sql)
        If Me.Database.LastQuery.Successful Then
            If Me._Id = 0 Then
                Me.Open(Me.PartNo)
                RaiseEvent NewItemMasterCreated()
            End If
        Else
            Throw New Exception(Me.Database.LastQuery.ErrorMsg)
        End If
    End Sub

    Public Sub SaveAlternatePricing(ByVal Table As DataTable)
        If Me.PartNo.Length > 0 And Me._Id > 0 Then
            If Table.Rows.Count > 0 Then
                ' Make list
                Dim List As String = ""
                For Each Row As DataRow In Table.Rows
                    List &= Row.Item("rate_table_id") & ", "
                Next
                List = List.Substring(0, List.LastIndexOf(","))
                ' Delete those not on list
                Dim Sql As String = ""
                Sql &= "DELETE FROM rate_table_item WHERE part_no=" & Me.Database.Escape(Me.PartNo)
                Sql &= " AND rate_table_id NOT IN (" & List & ")"
                Me.Database.Execute(Sql)
                ' Loop through and save
                For Each Row As DataRow In Table.Rows
                    If Row.RowState = DataRowState.Added Then
                        Sql = "INSERT INTO rate_table_item (rate_table_id, part_no, price,"
                        Sql &= " date_last_updated, last_updated_by)"
                        Sql &= " VALUES ("
                        Sql &= Me.Database.Escape(Row.Item("rate_table_id")) & ", "
                        Sql &= Me.Database.Escape(Me.PartNo) & ", "
                        Sql &= Me.Database.Escape(Row.Item("price")) & ", "
                        Sql &= Me.Database.Escape(Now) & ", "
                        Sql &= Me.Database.Escape(Me.LastUpdatedBy)
                        Sql &= ")"
                        Me.Database.Execute(Sql)
                        If Not Me.Database.LastQuery.Successful Then
                            Dim Err As String = Me.Database.LastQuery.ErrorMsg
                        End If
                    ElseIf Row.RowState = DataRowState.Modified Then
                        Sql = "UPDATE rate_table_item SET "
                        Sql &= " rate_table_id=" & Me.Database.Escape(Row.Item("rate_table_id")) & ", "
                        Sql &= " part_no=" & Me.Database.Escape(Me.PartNo) & ", "
                        Sql &= " price=" & Me.Database.Escape(Row.Item("price")) & ", "
                        Sql &= " date_last_updated=" & Me.Database.Escape(Now) & ", "
                        Sql &= " last_updated_by=" & Me.Database.Escape(Me.LastUpdatedBy)
                        Sql &= " WHERE id=" & Row.Item("id")
                        Me.Database.Execute(Sql)
                    End If
                Next
            Else
                Me.Database.Execute("DELETE FROM rate_table_item WHERE part_no=" & Me.Database.Escape(Me.PartNo))
            End If
        End If
    End Sub

    Public Function PurcahseOrders() As DataTable
        Dim Sql As String = "SELECT po_no, quantity, options, unit_price, description, discount, net_price"
        Sql &= " FROM purchase_order_item WHERE part_no=" & Me.Database.Escape(Me.PartNo)
        Sql &= " ORDER BY id DESC"
        Return Me.Database.GetAll(Sql)
    End Function

    Public Function ItemTypes() As DataTable
        Return Me.Database.GetAll("SELECT * FROM item_type ORDER BY sort, name")
    End Function

    Public Function IncomeAccounts() As DataTable
        Return Me.Database.GetAll("SELECT account_no, account_no + ' - ' + [description] AS name FROM gl_account WHERE active=1 AND account_type_id=1 ORDER BY account_no")
    End Function

    Public Function CostAccounts() As DataTable
        Return Me.Database.GetAll("SELECT account_no, account_no + ' - ' + [description] AS name FROM gl_account WHERE active=1 AND (account_type_id=3 OR account_type_id=6 OR account_type_id=2) ORDER BY account_no")
    End Function

    Public Function AssetAccounts() As DataTable
        Return Me.Database.GetAll("SELECT account_no, account_no + ' - ' + [description] AS name FROM gl_account WHERE active=1 AND (account_type_id=3 OR account_type_id=5 OR account_type_id=8) ORDER BY account_no")
    End Function

    Public Function RentalAccounts() As DataTable
        Return Me.Database.GetAll("SELECT account_no, account_no + ' - ' + [description] AS name FROM gl_account WHERE active=1 AND account_type_id=1 ORDER BY account_no")
    End Function

    Public Function RateTables() As DataTable
        Return Me.Database.GetAll("SELECT id, name FROM rate_table ORDER BY name, id")
    End Function

    Public Function RateTableItems() As DataTable
        If Me.PartNo.Length > 0 Then
            Dim Sql As String = ""
            Sql &= "SELECT rti.id, rti.rate_table_id, rti.price, rt.name"
            Sql &= " FROM rate_table_item rti"
            Sql &= " INNER JOIN rate_table rt ON rti.rate_table_id=rt.id"
            Sql &= " WHERE rti.part_no=" & Me.Database.Escape(Me.PartNo)
            Return Me.Database.GetAll(Sql)
        Else
            Return New DataTable
        End If
    End Function

    Public Function CheckQuantity(ByVal StationId As Integer) As Integer
        If Not Me.ItemType = 1 Then
            Throw New Exception("This part number is not for an inventory part.")
        ElseIf Me.Prime Then
            Return 0    ' Make this check equipment table
        Else
            Dim sql As String = "SELECT SUM(quantity) FROM item_inventory"
            sql &= " WHERE part_no=@part_no AND station_id=@station_id"
            sql &= " GROUP BY part_no"
            sql = sql.Replace("@part_no", Me.Database.Escape(Me.PartNo))
            sql = sql.Replace("@station_id", Me.Database.Escape(StationId))
            Try
                Dim Amount As Integer = Me.Database.GetOne(sql)
                Return Amount
            Catch
                Return 0
            End Try
        End If
    End Function

    Public Sub MoveNonPrimeInventory(ByVal FromStation As Integer, ByVal ToStation As Integer, ByVal Amount As Double, ByVal ReasonCode As InventoryChangeReason, ByVal ReasonMsg As String)
        Dim FromAmount As Integer = Me.CheckQuantity(FromStation)
        Dim ToAmount As Integer = Me.CheckQuantity(ToStation)
        If FromAmount >= Amount Then
            Dim Sql As String = ""
            ' Get to
            Dim toid As Integer = Me.Database.GetOne("SELECT id FROM item_inventory WHERE part_no=" & _
                Me.Database.Escape(Me.PartNo) & " AND station_id=" & ToStation)
            ' Substract from from
            Me.Database.Execute("UPDATE item_inventory SET quantity=quantity-" & Amount & _
                " WHERE part_no=" & Me.Database.Escape(Me.PartNo) & " AND " & _
                " station_id=" & FromStation)
            ' Add to to 
            If toid > 0 Then
                Me.Database.Execute("UPDATE item_inventory SET quantity=quantity+" & Amount & _
                    " WHERE id=" & toid)
            Else
                Me.Database.Execute("INSERT INTO item_inventory (part_no, station_id, quantity)" & _
                    " VALUES (" & _
                    Me.Database.Escape(Me.PartNo) & " , " & ToStation & ", " & _
                    Me.Database.Escape(Amount) & ")")
            End If
            ' Log history
            Sql = "INSERT INTO item_inventory_adjustment (part_no, from_station, to_station, from_quantity, to_quantity, "
            Sql &= " reason_id, note, created_by, date_created, amount)"
            Sql &= " VALUES (@part_no, @from, @to, @from_qty, @to_qty, @reason_id, @notes, @by, @when, @amount)"
            Sql = Sql.Replace("@part_no", Me.Database.Escape(Me.PartNo))
            Sql = Sql.Replace("@to_qty", ToAmount + Amount)
            Sql = Sql.Replace("@from_qty", FromAmount - Amount)
            Sql = Sql.Replace("@to", ToStation)
            Sql = Sql.Replace("@from", FromStation)
            Sql = Sql.Replace("@reason_id", CInt(ReasonCode))
            Sql = Sql.Replace("@notes", Me.Database.Escape(ReasonMsg))
            Sql = Sql.Replace("@by", Me.Database.Escape(Me.LastUpdatedBy))
            Sql = Sql.Replace("@when", Me.Database.Escape(Now.ToString))
            Sql = Sql.Replace("@amount", Amount)
            Me.Database.Execute(Sql)
            If Not Me.Database.LastQuery.Successful Then
                Throw New Exception(Me.Database.LastQuery.ErrorMsg)
            End If
        Else
            Throw New Exception("There is not sufficient quantity at the from location.")
        End If
    End Sub

    Public Sub AdjustNonPrimeInventory(ByVal StationId As Integer, ByVal Adjustment As Integer, ByVal ReasonCode As InventoryChangeReason, ByVal ReasonMsg As String)
        If Not Me.Prime And Me.ItemType = 1 Then
            ' Get inventory id 
            Dim Sql As String = "SELECT id, quantity FROM item_inventory WHERE part_no=" & Me.Database.Escape(Me.PartNo) & " AND station_id=" & StationId
            Dim Row As DataRow = Me.Database.GetRow(Sql)
            If Me.Database.LastQuery.RowsReturned = 1 Then
                ' Take from quantity
                Sql = "UPDATE item_inventory SET quantity=quantity+@adjust WHERE id=@id"
                Sql = Sql.Replace("@id", Row.Item("id"))
            Else
                ' Insert into table
                Sql = "INSERT INTO item_inventory (part_no, station_id, quantity) VALUES (@part_no, @station_id, @adjust)"
            End If
            Sql = Sql.Replace("@adjust", Adjustment)
            Sql = Sql.Replace("@part_no", Me.Database.Escape(Me.PartNo))
            Sql = Sql.Replace("@station_id", StationId)
            Me.Database.Execute(Sql)
            If Not Me.Database.LastQuery.Successful Then
                Throw New Exception(Me.Database.LastQuery.ErrorMsg)
            End If
            ' Log history
            Sql = "INSERT INTO item_inventory_adjustment (part_no, from_station, to_station, from_quantity, to_quantity, "
            Sql &= " reason_id, note, created_by, date_created, amount)"
            Sql &= " VALUES (@part_no, @from, @to, @from_qty, @to_qty, @reason_id, @notes, @by, @when, @amount)"
            Sql = Sql.Replace("@part_no", Me.Database.Escape(Me.PartNo))
            Sql = Sql.Replace("@to_qty", 0)
            Try
                Sql = Sql.Replace("@from_qty", Row.Item("quantity") + Adjustment)
            Catch
                Sql = Sql.Replace("@from_qty", Adjustment)
            End Try
            Sql = Sql.Replace("@to", 0)
            Sql = Sql.Replace("@from", StationId)
            Sql = Sql.Replace("@reason_id", ReasonCode)
            Sql = Sql.Replace("@notes", Me.Database.Escape(ReasonMsg))
            Sql = Sql.Replace("@by", Me.Database.Escape(Me.LastUpdatedBy))
            Sql = Sql.Replace("@when", Me.Database.Escape(Now.ToString))
            Sql = Sql.Replace("@amount", Adjustment)
            Me.Database.Execute(Sql)
            If Not Me.Database.LastQuery.Successful Then
                Throw New Exception(Me.Database.LastQuery.ErrorMsg)
            End If
        Else
            Throw New Exception("The quantity for part #" & Me.PartNo & " can not be adjusted in this manner.")
        End If
    End Sub

    Public Function GetPrice(ByVal RateTableId As Integer) As Double
        If RateTableId > 0 Then
            ' Open rate table
            Dim RateTable As New cRateTable(Me.Database)
            RateTable.Open(RateTableId)
            ' Look if the price was overridden for this rate table
            Dim OverridePrice As Double = RateTable.GetPrice(Me.PartNo, True)
            ' Decide what to do
            If OverridePrice <> Nothing Then
                ' If price was overridden in the rate table, return that value
                Return OverridePrice
            Else
                ' Otherwise do calculation
                If RateTable.SalesPriceFormula.Length > 0 Then
                    ' If there is a formula in the rate table
                    Dim Formula As New MyCore.mcCalc
                    Formula.AddVariable("list", Me.ListPrice)
                    Formula.AddVariable("cost", Me.Cost)
                    Formula.AddVariable("freight", Me.Freight)
                    Formula.AddVariable("discount", Me.Cost / Me.ListPrice)
                    Formula.AddVariable("type", Me.ItemType)
                    Try
                        Return Formula.Eval(RateTable.SalesPriceFormula)
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                        Return Me.ListPrice
                    End Try
                Else
                    ' Otherwise return our list price
                    Return Me.ListPrice
                End If
            End If
        Else
            Try
                Return Me.ListPrice
            Catch
                Return 0
            End Try
        End If
    End Function

    Public Function GetChangeHistory(ByVal StockRoomId As Integer) As DataTable
        If Me.PartNo.Length > 0 And Not Me.Prime And Me.ItemType = 1 Then
            Dim Sql As String = "SELECT a.*, r.name AS reason, s1.name AS from_name, s2.name AS to_name"
            Sql &= " FROM item_inventory_adjustment a"
            Sql &= " LEFT OUTER JOIN item_change_reason r ON a.reason_id=r.id"
            Sql &= " LEFT OUTER JOIN station s1 ON a.from_station=s1.id"
            Sql &= " LEFT OUTER JOIN station s2 ON a.to_station=s2.id"
            Sql &= " WHERE part_no=" & Me.Database.Escape(Me.PartNo)
            Sql &= " AND (to_station=" & StockRoomId
            Sql &= " OR from_station=" & StockRoomId & ")"
            Sql &= " ORDER BY a.date_created DESC"
            Return Me.Database.GetAll(Sql)
        Else
            Throw New Exception("Cannot call this function on this type of part.")
            Return Nothing
        End If
    End Function

    Public Function GetXRefs() As DataTable
        Return Me.Database.GetAll("SELECT old_part_no FROM part_no_xref WHERE new_part_no=" & Me.Database.Escape(Me.PartNo))
    End Function

    Public Function GetBins() As DataTable
        Return Me.Database.GetAll("SELECT id, part_no, station_id, name FROM item_master_bin")
    End Function

    Public Sub AddBin(ByVal StationID As Integer, ByVal Name As String)
        Me.Database.Execute("INSERT INTO item_master_bin (part_no, station_id, name)" & _
                            " VALUES (" & Me.Database.Escape(Me.PartNo) & ", " & _
                            Me.Database.Escape(StationID) & ", " & Me.Database.Escape(Name) & ")")
    End Sub

    Public Sub EditBin(ByVal Id As Integer, ByVal StationID As Integer, ByVal Name As String)
        Me.Database.Execute("UPDATE item_master_bin SET station_id=" & Me.Database.Escape(StationID) & _
                            ", name=" & Me.Database.Escape(Name) & _
                            " WHERE id=" & Id)
    End Sub

End Class
