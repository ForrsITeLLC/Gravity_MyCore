Imports MyCore.Data

Public Class cContact

    Public FirstName As String = ""
    Public LastName As String = ""
    Dim _Company As String = ""
    Public LeadSource As Integer = 0
    Public Department As String = ""
    Public Title As String = ""
    Public ReportsTo As Integer = 0
    Public Note As String = ""
    Public BusinessPhone As String = ""
    Public CellPhone As String = ""
    Public Fax As String = ""
    Public Email As String = ""
    Public Birthday As Date = Nothing
    Public DoNotEmail As Boolean = False
    Public DoNotCall As Boolean = False
    Public DoNotMail As Boolean = False
    Public Id As Integer = Nothing
    Public Address1 As String = ""
    Public Address2 As String = ""
    Public City As String = ""
    Public State As String = ""
    Public ZipCode As String = ""
    Public Salutation As String = ""
    Public Type As ContactType = ContactType.Primary

    Dim Database As MyCore.Data.EasySql

    Dim _ReportsTo As DataTable
    Dim _LeadSource As DataTable

    Public Event Reload()
    Public Event Saved(ByVal Contact As cContact)

    Public Enum ContactType
        Archived = 0
        Primary = 1
        Secondary = 2
    End Enum

    Public ReadOnly Property DisplayName() As String
        Get
            If Me.FirstName.Length > 0 And Me.LastName.Length > 0 Then
                Return Me.LastName & ", " & Me.FirstName
            ElseIf Me.FirstName.Length > 0 Then
                Return Me.FirstName
            ElseIf Me.LastName.Length > 0 Then
                Return Me.LastName
            Else
                Return "Unknown"
            End If
        End Get
    End Property

    Public ReadOnly Property DisplayNameFirstLast() As String
        Get
            If Me.FirstName.Length > 0 And Me.LastName.Length > 0 Then
                Return Me.FirstName & " " & Me.LastName
            ElseIf Me.FirstName.Length > 0 Then
                Return Me.FirstName
            ElseIf Me.LastName.Length > 0 Then
                Return Me.LastName
            Else
                Return "Unknown"
            End If
        End Get
    End Property

    Public ReadOnly Property ReportsToList() As DataTable
        Get
            Return Me._ReportsTo
        End Get
    End Property

    Public ReadOnly Property LeadSourceList() As DataTable
        Get
            Return Me._LeadSource
        End Get
    End Property

    Public Property Company() As String
        Get
            Return Me._Company
        End Get
        Set(ByVal Value As String)
            If Not Value = Me._Company Then
                If Value.Length > 0 Then
                    Me._Company = Value
                    Me.ReportsTo = 0
                    Me.PopulateReportsTo()
                End If
            End If
        End Set
    End Property


    Public Sub New(ByRef db As MyCore.Data.EasySql)
        Me.Database = db
        Me._LeadSource = Me.Database.GetAll("SELECT id, name FROM lead_source ORDER BY sort")
    End Sub

    Private Sub PopulateReportsTo()
        Dim Sql As String = "SELECT *, cnt_last + ', ' + cnt_first AS display_name"
        Sql &= " FROM CONTACTS WHERE cnt_no= " & Me.Database.Escape(Me._Company)
        If Me.Id > 0 Then
            Sql &= " AND cnt_id <> " & Me.Id
        End If
        Me._ReportsTo = Me.Database.GetAll(Sql)
    End Sub

    Public Sub Open(ByVal Id As Integer)
        Dim Sql As String = "SELECT"
        Sql &= " CONTACTS.cnt_id, CONTACTS.cnt_no, CONTACTS.cnt_last, CONTACTS.cnt_first,"
        Sql &= "  ISNULL(CONTACTS.cnt_title, '') AS cnt_title,"
        Sql &= " ISNULL(CONTACTS.cnt_dept, '') AS cnt_dept,"
        Sql &= " ISNULL(CONTACTS.cnt_mail, '') AS cnt_mail,"
        Sql &= " ISNULL(CONTACTS.cnt_phone, '') AS cnt_phone,"
        Sql &= " ISNULL(CONTACTS.cnt_fax, '') AS cnt_fax,"
        Sql &= " ISNULL(CONTACTS.cnt_memo, '') AS cnt_memo,"
        Sql &= " CONTACTS.cnt_chngt,"
        Sql &= " CONTACTS.cnt_user, CONTACTS.cnt_type,"
        Sql &= " ISNULL(CONTACTS.cnt_cell, '') AS cnt_cell,"
        'Sql &= " ADDRESS.cst_name,"
        'Sql &= " ADDRESS.cst_city,"
        'Sql &= " ADDRESS.cst_state,"
        Sql &= " (CONTACTS.cnt_last + ', ' + CONTACTS.cnt_first) AS display_name,"
        Sql &= " cnt_bday, lead_source_id, reports_to_id,"
        Sql &= " no_email, no_call, no_mail,"
        Sql &= " ISNULL(CONTACTS.address_line1, '') AS address_line1,"
        Sql &= " ISNULL(CONTACTS.address_line2, '') AS address_line2,"
        Sql &= " ISNULL(CONTACTS.city, '') AS city,"
        Sql &= " ISNULL(CONTACTS.state, '') AS state,"
        Sql &= " ISNULL(CONTACTS.zip, '') AS zip,"
        Sql &= " ISNULL(CONTACTS.salutation, '') AS salutation"
        Sql &= " FROM CONTACTS"
        'Sql &= " INNER JOIN ADDRESS ON CONTACTS.cnt_no = ADDRESS.cst_no"
        Sql &= " WHERE CONTACTS.cnt_id = " & Id
        Dim Row As DataRow = Me.Database.GetRow(Sql)
        If Me.Database.LastQuery.Successful Then
            Me.Id = Id
            Me.FirstName = IIf(Row.Item("cnt_first") Is DBNull.Value, "", Row.Item("cnt_first"))
            Me.LastName = IIf(Row.Item("cnt_last") Is DBNull.Value, "", Row.Item("cnt_last"))
            Me.Company = Row.Item("cnt_no")
            Me.Title = Row.Item("cnt_title")
            Me.Department = Row.Item("cnt_dept")
            Me.Email = Row.Item("cnt_mail")
            Me.BusinessPhone = Row.Item("cnt_phone")
            Me.Fax = Row.Item("cnt_fax")
            Me.CellPhone = Row.Item("cnt_cell")
            Me.Birthday = IIf(Row.Item("cnt_bday") Is DBNull.Value, Nothing, Row.Item("cnt_bday"))
            Me.ReportsTo = Row.Item("reports_to_id")
            Me.LeadSource = Row.Item("lead_source_id")
            Me.Note = Row.Item("cnt_memo")
            Me.Address1 = Row.Item("address_line1")
            Me.Address2 = Row.Item("address_line2")
            Me.City = Row.Item("city")
            Me.State = Row.Item("state")
            Me.ZipCode = Row.Item("zip")
            Me.Salutation = Row.Item("salutation")
            Me.DoNotEmail = Me.IsNull(Row.Item("no_email"), False)
            Me.DoNotCall = Me.IsNull(Row.Item("no_call"), False)
            Me.DoNotMail = Me.IsNull(Row.Item("no_mail"), False)
            RaiseEvent Reload()
        Else
            Dim Err As String = Me.Database.LastQuery.ErrorMsg
            MsgBox(Err)
        End If
    End Sub

    Public Function Save() As Boolean
        Dim Sql As String = ""
        If Me.Id = Nothing Then
            Sql = "INSERT INTO CONTACTS"
            Sql &= " (cnt_first, cnt_last, cnt_no, cnt_dept, cnt_title, cnt_mail, cnt_phone, cnt_fax, "
            Sql &= " cnt_memo, cnt_cell, cnt_bday, cnt_chngt, reports_to_id, lead_source_id,"
            Sql &= " no_email, no_call, no_mail,"
            Sql &= " address_line1, address_line2, city, state, zip, salutation)"
            Sql &= " VALUES ( @first_name, @last_name, @customer_no, @department, @title,"
            Sql &= " @email, @phone, @fax, @note, @cell, @birthday, " & Me.Database.Timestamp
            Sql &= ", @reports_to_id, @lead_source_id"
            Sql &= ", @no_email, @no_call, @no_mail, @address_line1, @address_line2"
            Sql &= ", @city, @state, @zip, @salutation"
            Sql &= ")"
        Else
            Sql = "UPDATE CONTACTS SET"
            Sql &= " cnt_first = @first_name,"
            Sql &= " cnt_last = @last_name,"
            Sql &= " cnt_no = @customer_no,"
            Sql &= " cnt_phone = @phone,"
            Sql &= " cnt_fax = @fax,"
            Sql &= " cnt_cell = @cell,"
            Sql &= " cnt_mail = @email,"
            Sql &= " cnt_title = @title,"
            Sql &= " cnt_dept = @department,"
            Sql &= " cnt_memo = @note,"
            Sql &= " cnt_bday = @birthday,"
            Sql &= " cnt_chngt = " & Me.Database.Timestamp & ","
            Sql &= " reports_to_id = @reports_to_id,"
            Sql &= " lead_source_id = @lead_source_id,"
            Sql &= " no_email = @no_email,"
            Sql &= " no_call = @no_call,"
            Sql &= " no_mail = @no_mail,"
            Sql &= " address_line1 = @address_line1,"
            Sql &= " address_line2 = @address_line2,"
            Sql &= " city = @city,"
            Sql &= " state = @state,"
            Sql &= " zip = @zip,"
            Sql &= " salutation = @salutation"
            Sql &= " WHERE"
            Sql &= " cnt_id = @id"
        End If
        Sql = Sql.Replace("@first_name", Me.Database.Escape(Me.FirstName))
        Sql = Sql.Replace("@last_name", Me.Database.Escape(Me.LastName))
        Sql = Sql.Replace("@customer_no", Me.Database.Escape(Me.Company))
        Sql = Sql.Replace("@phone", Me.Database.Escape(Me.BusinessPhone))
        Sql = Sql.Replace("@fax", Me.Database.Escape(Me.Fax))
        Sql = Sql.Replace("@cell", Me.Database.Escape(Me.CellPhone))
        Sql = Sql.Replace("@email", Me.Database.Escape(Me.Email))
        Sql = Sql.Replace("@title", Me.Database.Escape(Me.Title))
        Sql = Sql.Replace("@department", Me.Database.Escape(Me.Department))
        Sql = Sql.Replace("@note", Me.Database.Escape(Me.Note))
        Sql = Sql.Replace("@birthday", Me.Database.Escape(IIf(Me.Birthday = Nothing, DBNull.Value, Me.Birthday)))
        Sql = Sql.Replace("@reports_to_id", Me.Database.Escape(Me.ReportsTo))
        Sql = Sql.Replace("@lead_source_id", Me.Database.Escape(Me.LeadSource))
        Sql = Sql.Replace("@id", Me.Id)
        Sql = Sql.Replace("@no_email", Me.Database.Escape(Me.DoNotEmail))
        Sql = Sql.Replace("@no_call", Me.Database.Escape(Me.DoNotCall))
        Sql = Sql.Replace("@no_mail", Me.Database.Escape(Me.DoNotMail))
        Sql = Sql.Replace("@address_line1", Me.Database.Escape(Me.Address1))
        Sql = Sql.Replace("@address_line2", Me.Database.Escape(Me.Address2))
        Sql = Sql.Replace("@city", Me.Database.Escape(Me.City))
        Sql = Sql.Replace("@state", Me.Database.Escape(Me.State))
        Sql = Sql.Replace("@zip", Me.Database.Escape(Me.ZipCode))
        Sql = Sql.Replace("@salutation", Me.Database.Escape(Me.Salutation))

        If Me.Id = Nothing Then
            Me.Database.InsertAndReturnId(Sql)
        Else
            Me.Database.Execute(Sql)
        End If

        If Not Me.Database.LastQuery.Successful Then
            Return False
        Else
            If Me.Id = Nothing Then
                Me.Id = Me.Database.LastQuery.InsertId
            End If
            RaiseEvent Saved(Me)
            Return True
        End If
    End Function

    Public Function GetInteractions() As DataTable
        Dim Sql As String = "SELECT journal.contact_name, journal.contact_id, journal.customer_no, journal.id AS journal_id, journal.entry_type_id, journal.subject, journal.memo,"
        Sql &= " journal.created_by, journal.created_date, journal.touch_date, journal.touch_by, journal_entry_type.name AS entry_type, "
        Sql &= " CASE journal.department WHEN 1 THEN 'Sales' WHEN 2 THEN 'Service' WHEN 3 THEN 'Collections' WHEN 4 THEN 'Rental' ELSE 'Purchase' END AS department_name,"
        Sql &= " CASE journal.initiator WHEN 1 THEN 'Us' ELSE 'Them' END AS initiator_name,"
        Sql &= " journal.department, journal.initiator"
        Sql &= " FROM journal"
        Sql &= " INNER JOIN journal_entry_type ON journal.entry_type_id = journal_entry_type.id "
        Sql &= " WHERE journal.contact_id=" & Me.Database.Escape(Me.Id)
        Sql &= " ORDER BY journal.touch_date DESC"
        Return Me.Database.GetAll(Sql)
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
        Sql &= " WHERE so.contact_id=" & Me.Database.Escape(Me.Id)
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
        Sql &= " WHERE so.contact_id=" & Me.Database.Escape(Me.Id)
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
        Sql &= " WHERE ro.contact_id=" & Me.Database.Escape(Me.Id)
        Sql &= " ORDER BY ro.date_created DESC"
        Return Me.Database.GetAll(Sql)
    End Function

    Public Function GetPurchaseOrders(Optional ByVal IncludeRMAs As Boolean = False) As DataTable
        Dim Sql As String = "SELECT *,"
        Sql &= " CASE [rma] WHEN 0 THEN 'PO' ELSE 'RMA' END AS [type],"
        Sql &= " CASE WHEN (SELECT SUM(quantity) FROM purchase_order_item WHERE po_no=po.po_no AND received < quantity) > 0 THEN 'N' ELSE 'Y' END AS is_received"
        Sql &= " FROM purchase_order po WHERE contact_id=" & Me.Database.Escape(Me.Id)
        If Not IncludeRMAs Then
            Sql &= " AND rma=0"
        End If
        Sql &= " ORDER BY po.po_date DESC"
        Dim Table As DataTable = Me.Database.GetAll(Sql)
        If Me.Database.LastQuery.Successful Then
            Return Table
        Else
            Throw New Exception("Error getting purchase orders.  Error from database: " & Me.Database.LastQuery.ErrorMsg)
        End If
    End Function

    Private Function IsNull(ByVal Value As Object, Optional ByVal DefaultValue As Object = "") As Object
        If Value Is DBNull.Value Then
            Return DefaultValue
        Else
            Return Value
        End If
    End Function

End Class
