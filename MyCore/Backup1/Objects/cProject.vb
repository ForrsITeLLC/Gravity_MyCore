Public Class cProject

    Dim Database As MyCore.Data.EasySql
    Dim _Id As Integer = 0
    Dim _ErrorMsg As String
    Dim _CreatedBy As String
    Dim _DateCreated As Date

    Public Name As String = ""
    Public CustomerNo As String = ""
    Public ContactId As Integer = 0
    Public ContactName As String = ""
    Public ProjectedCost As Double = 0
    Public ProjectedProfit As Double = 0
    Public DateStart As Date = Nothing
    Public DateProjectedClose As Date = Nothing
    Public DateCompleted As Date = Nothing
    Public PercentCompleted As Integer = 0
    Public Notes As String = ""
    Public LeadId As Integer = 0

    Public LastUpdatedBy As String = ""

    Public ReadOnly Property CreatedBy() As String
        Get
            Return Me._CreatedBy
        End Get
    End Property

    Public Event Saved(ByVal Project As cProject)

    Public ReadOnly Property Id() As Integer
        Get
            Return Me._Id
        End Get
    End Property

    Public ReadOnly Property ErrorMsg() As String
        Get
            Return Me._ErrorMsg
        End Get
    End Property


    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
    End Sub

    Public Function Open(ByVal Id As Integer) As Boolean
        Dim Row As DataRow = Me.Database.GetRow("SELECT * FROM project WHERE id=" & Id)
        If Me.Database.LastQuery.RowsReturned = 1 Then
            Me._Id = Id
            Me.Name = Row.Item("name")
            Me.CustomerNo = Row.Item("customer_no")
            Me.ContactId = Row.Item("contact_id")
            Me.ContactName = Row.Item("contact_name")
            Me.ProjectedCost = Row.Item("projected_cost")
            Me.ProjectedProfit = Row.Item("projected_profit")
            Me.DateStart = Row.Item("date_start")
            Me.DateProjectedClose = Row.Item("date_projected_close")
            If Row.Item("date_completed") Is DBNull.Value Then
                Me.DateCompleted = Nothing
            Else
                Me.DateCompleted = Row.Item("date_completed")
            End If
            Me.PercentCompleted = Row.Item("percent_completed")
            Me.Notes = Row.Item("notes")
            Me.LeadId = Row.Item("lead_id")
            Me._CreatedBy = Row.Item("created_by")
            Me._DateCreated = Row.Item("date_created")
            Me.LastUpdatedBy = Row.Item("last_updated_by")
            Return True
        Else
            If Me.Database.LastQuery.Successful Then
                Me._ErrorMsg = "Project not found."
            Else
                Me._ErrorMsg = Me.Database.LastQuery.ErrorMsg
            End If
            Return False
        End If
    End Function

    Public Function Save() As Boolean
        Dim Sql As String = ""
        If Me._Id = 0 Then
            Sql &= "INSERT INTO project "
            Sql &= " (name, customer_no, contact_id, contact_name, projected_cost, projected_profit, date_start,"
            Sql &= " date_projected_close, date_completed, percent_completed, lead_id, notes, "
            Sql &= " date_created, created_by, date_last_updated, last_updated_by)"
            Sql &= " VALUES ("
            Sql &= " @name, @customer_no, @contact_id, @contact_name, @projected_cost, @projected_profit,"
            Sql &= " @date_start, @date_projected_close, @date_completed, @percent_completed, @lead_id, @notes,"
            Sql &= " " & Me.Database.Timestamp & ", @user, " & Me.Database.Timestamp & ", @user)"
        Else
            Sql &= "UPDATE project SET"
            Sql &= " name=@name, customer_no=@customer_no, contact_id=@contact_id, "
            Sql &= " contact_name=@contact_name, projected_cost=@projected_cost, "
            Sql &= " projected_profit=@projected_profit, date_start=@date_start,"
            Sql &= " date_projected_close=@date_projected_close, date_completed=@date_completed, "
            Sql &= " percent_completed=@percent_completed, lead_id=@lead_id, "
            Sql &= " notes=@notes, date_last_updated=" & Me.Database.Timestamp & ", last_updated_by=@user"
            Sql &= " WHERE id=@id"
        End If
        Sql = Sql.Replace("@name", Me.Database.Escape(Me.Name))
        Sql = Sql.Replace("@customer_no", Me.Database.Escape(Me.CustomerNo))
        Sql = Sql.Replace("@contact_id", Me.Database.Escape(Me.ContactId))
        Sql = Sql.Replace("@contact_name", Me.Database.Escape(Me.ContactName))
        Sql = Sql.Replace("@projected_cost", Me.Database.Escape(Me.ProjectedCost))
        Sql = Sql.Replace("@projected_profit", Me.Database.Escape(Me.ProjectedProfit))
        Sql = Sql.Replace("@date_start", Me.Database.Escape(Me.DateStart))
        Sql = Sql.Replace("@date_projected_close", Me.Database.Escape(Me.DateProjectedClose))
        If Me.DateCompleted <> Nothing And Me.PercentCompleted = 100 Then
            Sql = Sql.Replace("@date_completed", Me.Database.Escape(Me.DateCompleted))
        Else
            Sql = Sql.Replace("@date_completed", Me.Database.Escape(DBNull.Value))
        End If
        Sql = Sql.Replace("@percent_completed", Me.Database.Escape(Me.PercentCompleted))
        Sql = Sql.Replace("@notes", Me.Database.Escape(Me.Notes))
        Sql = Sql.Replace("@lead_id", Me.Database.Escape(Me.LeadId))
        Sql = Sql.Replace("@user", Me.Database.Escape(Me.LastUpdatedBy))
        If Me._Id > 0 Then
            Sql = Sql.Replace("@id", Me.Database.Escape(Me._Id))
            Me.Database.Execute(Sql)
        Else
            Me.Database.InsertAndReturnId(Sql)
        End If
        If Me.Database.LastQuery.Successful Then
            RaiseEvent Saved(Me)
            If Me._Id = 0 Then
                Me._Id = Me.Database.LastQuery.InsertId
            End If
            Return True
        Else
            Me._ErrorMsg = Me.Database.LastQuery.ErrorMsg
            Return False
        End If
    End Function

    Public Function ServiceOrders() As DataTable
        Dim Sql As String = "SELECT DISTINCT so.[id], "
        Sql &= " l.cst_name AS company, l.cst_no AS customer_no,"
        Sql &= " a.cst_name AS bill_to_company, a.cst_no AS bill_to_customer_no,"
        Sql &= " so.date_created, so.date_due, so.assigned_to, so.date_completed,"
        Sql &= " (SELECT TOP 1 date_start FROM schedule WHERE reference_id=so.id AND deleted=0 ORDER BY date_start ASC) AS date_scheduled,"
        Sql &= " so.contact_name  AS [contact_display_name],"
        Sql &= " so.contact_name, so.caller_name, so.approved_name,"
        Sql &= " work_location = CASE so.shop_work WHEN 1 THEN 'Shop' ELSE 'On Site' END,"
        Sql &= " type = sot.name,"
        Sql &= " so.[work], so.created_by,"
        Sql &= " l.cst_zip AS location_zip, z.lon, z.lat,"
        Sql &= " stage = CASE "
        Sql &= " WHEN so.date_completed IS NOT NULL THEN 'Completed'"
        Sql &= " WHEN so.date_scheduled IS NOT NULL THEN 'Scheduled'"
        Sql &= " ELSE 'Received'"
        Sql &= " End"
        Sql &= " FROM service_order so"
        Sql &= " INNER JOIN ADDRESS a ON so.customer_no=a.cst_no"
        Sql &= " INNER JOIN ADDRESS l ON so.location_id=l.cst_no"
        Sql &= " INNER JOIN service_order_type sot ON sot.id=so.service_order_type"
        Sql &= " LEFT OUTER JOIN zips z ON l.cst_zip=z.zip"
        Sql &= " WHERE so.project_id=@project_id"
        Sql &= " ORDER BY so.date_created DESC"
        Sql = Sql.Replace("@project_id", Me.Database.Escape(Me.Id))
        Return Me.Database.GetAll(Sql)
    End Function

    Public Function SalesOrders() As DataTable
        Dim Sql As String = "SELECT so.*, "
        Sql &= " s.cst_name AS ship_to_name, s.cst_city + ', ' + s.cst_state AS ship_to_city,"
        Sql &= " b.cst_name AS bill_to_name, b.cst_city + ', ' + b.cst_state AS bill_to_city"
        Sql &= " FROM sales_order so"
        Sql &= " LEFT JOIN ADDRESS s ON so.ship_to=s.cst_no"
        Sql &= " LEFT JOIN ADDRESS b ON so.bill_to=b.cst_no"
        Sql &= " WHERE so.project_id=@project_id"
        Sql = Sql.Replace("@project_id", Me.Database.Escape(Me.Id))
        Return Me.Database.GetAll(Sql)
    End Function

    Public Function RentalOrders() As DataTable
        Dim Sql As String = "SELECT ro.*, "
        Sql &= " s.cst_name AS ship_to_name, s.cst_city + ', ' + s.cst_state AS ship_to_city,"
        Sql &= " b.cst_name AS bill_to_name, b.cst_city + ', ' + b.cst_state AS bill_to_city"
        Sql &= " FROM rental_order ro"
        Sql &= " LEFT JOIN ADDRESS s ON ro.ship_to_no=s.cst_no"
        Sql &= " LEFT JOIN ADDRESS b ON ro.bill_to_no=b.cst_no"
        Sql &= " WHERE ro.project_id=@project_id"
        Sql = Sql.Replace("@project_id", Me.Database.Escape(Me.Id))
        Return Me.Database.GetAll(Sql)
    End Function

    Public Function PurchaseOrders() As DataTable
        Dim Sql As String = "SELECT po.[id], po.po_no, po.vendor_no, c.cst_name AS vendor_name, "
        Sql &= " po.po_date, po.date_ordered, po.date_planned_ship, po.date_expected,"
        Sql &= " ((SELECT SUM(quantity*unit_price) FROM purchase_order_item WHERE po_no=po.po_no) + shipping_charge + tax) AS total_price,"
        Sql &= " po.office, po.requested_by, "
        Sql &= " received = CAST((CASE (SELECT COUNT(id) FROM purchase_order_item poi WHERE poi.po_no=po.po_no AND quantity>received) WHEN 0  THEN 1 ELSE 0 END) AS bit)"
        Sql &= " FROM purchase_order po"
        Sql &= " LEFT JOIN ADDRESS c ON po.vendor_no=c.cst_no"
        Sql &= " WHERE"
        Sql &= " po.type=2 AND po.our_order_no in (SELECT id FROM sales_order WHERE project_id=@project_id)"
        Sql &= " OR po.type=4 AND po.our_order_no in (SELECT id FROM rental_order WHERE project_id=@project_id)"
        Sql &= " OR po.type=3 AND po.our_order_no in (SELECT id FROM service_order WHERE project_id=@project_id)"
        Sql = Sql.Replace("@project_id", Me.Database.Escape(Me.Id))
        Return Me.Database.GetAll(Sql)
    End Function

    Public Function Invoices() As DataTable
        Dim Sql As String = "SELECT i.* FROM INVOICE i WHERE i.inv_ino IN (SELECT invoice_no FROM sales_order WHERE project_id=@project_id)"
        Sql &= " UNION ALL"
        Sql &= " SELECT i.* FROM INVOICE i WHERE i.inv_ino IN (SELECT invoice_no FROM rental_order WHERE project_id=@project_id)"
        Sql &= " UNION ALL"
        Sql &= " SELECT i.* FROM INVOICE i WHERE i.inv_ino IN  (SELECT invoice_id FROM service_order WHERE project_id=@project_id AND charge_to=0)"
        Sql = Sql.Replace("@project_id", Me.Database.Escape(Me.Id))
        Return Me.Database.GetAll(Sql)
    End Function


End Class
