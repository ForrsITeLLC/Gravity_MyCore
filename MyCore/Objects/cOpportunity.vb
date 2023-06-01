Public Class cOpportunity

    Dim _Id As Integer = 0

    Public Name As String = ""
    Public BillToNo As String = ""
    Public BillToName As String = ""
    Public ShipToNo As String = ""
    Public ShipToName As String = ""
    Public ContactId As Integer = 0
    Public ContactName As String = ""
    Public Type As Integer = 1
    Public LeadSourceId As Integer = 0
    Public Salesperson As String = ""
    Public EmployeeReferral As String = ""
    Public Office As String = ""
    Public DateExpectedClose As Date = Today
    Public DateActualClose As Date = Nothing
    Public StageId As Integer = 1
    Public Amount As Double = 0
    Public Notes As String = ""
    Public Probability As Integer = 0
    Public ReasonLost As String = ""
    Public DateCreated As DateTime = Now
    Public DateLastUpdated As DateTime = Now
    Public LastUpdatedBy As String = ""

    Public Database As MyCore.Data.EasySql

    Public Event Reload()
    Public Event Saved(ByVal Opp As cOpportunity)

    Public ReadOnly Property Id() As Integer
        Get
            Return Me._Id
        End Get
    End Property

    Public ReadOnly Property Offices() As DataTable
        Get
            Return Me.Database.GetAll("SELECT id, number, name, sort FROM office ORDER BY sort")
        End Get
    End Property

    Public ReadOnly Property Stages() As DataTable
        Get
            Return Me.Database.GetAll("SELECT id, name, probability, sort FROM lead_status ORDER BY sort, name")
        End Get
    End Property

    Public ReadOnly Property Sources() As DataTable
        Get
            Return Me.Database.GetAll("SELECT id, name, sort FROM lead_source ORDER BY sort, name")
        End Get
    End Property

    Public ReadOnly Property Employees() As DataTable
        Get
            Return Me.Database.GetAll("SELECT windows_user, last_name + ', ' + first_name AS display_name FROM employee WHERE deactivated=0 ORDER BY last_name, first_name")
        End Get
    End Property

    Public ReadOnly Property Quotes() As DataTable
        Get
            Return Me.Database.GetAll("SELECT * FROM quote WHERE opportunity_id=" & Me.Id)
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
            Sql &= " WHERE ref_no=" & Me.Database.Escape(Me.Id)
            Sql &= " AND department=" & CInt(cInteraction.ReferenceTypes.Opportunity)
            Sql &= " ORDER BY touch_date DESC"
            Return Me.Database.GetAll(Sql)
        End Get
    End Property

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
    End Sub

    Public Sub Open(ByVal Id As Integer)
        Dim Sql As String = ""
        Sql &= "SELECT opportunity.*, CONTACTS.cnt_last + ', ' + CONTACTS.cnt_first AS display_name,"
        Sql &= " b.cst_name AS bill_to_name, s.cst_name = ship_to_name"
        Sql &= " FROM opportunity"
        Sql &= " LEFT JOIN CONTACTS ON opportunity.contact_id=CONTACTS.cnt_id"
        Sql &= " LEFT JOIN ADDRESS b ON quote.bill_to_no=b.cst_no"
        Sql &= " LEFT JOIN ADDRESS s ON quote.ship_to_no=s.cst_no"
        Sql &= " WHERE opportunity.id = " & Id
        Dim Row As DataRow = Me.Database.GetRow(Sql)
        If Me.Database.LastQuery.Successful And Me.Database.LastQuery.RowsReturned = 1 Then
            Me.Amount = Row.Item("amount")
            Me.BillToNo = Row.Item("bill_to_no")
            Me.BillToName = Row.Item("bill_to_name")
            Me.ContactId = Row.Item("contact_id")
            Me.ContactName = Row.Item("display_name")
            Me.DateActualClose = IIf(Row.Item("date_actual_close") Is DBNull.Value, Nothing, Row.Item("date_actual_close"))
            Me.DateCreated = Row.Item("date_created")
            Me.DateExpectedClose = Row.Item("date_expected_close")
            Me.DateLastUpdated = Row.Item("date_last_updated")
            Me.EmployeeReferral = Row.Item("referral")
            Me._Id = Id
            Me.LastUpdatedBy = Row.Item("last_updated_by")
            Me.LeadSourceId = Row.Item("lead_source_id")
            Me.Name = Row.Item("name")
            Me.Notes = Row.Item("notes")
            Me.Office = Row.Item("office")
            Me.Probability = Row.Item("probability")
            Me.ReasonLost = Row.Item("reason_lost")
            Me.Salesperson = Row.Item("salesperson")
            Me.ShipToNo = Row.Item("ship_to_no")
            Me.ShipToName = Row.Item("ship_to_name")
            Me.Type = Row.Item("type")
            RaiseEvent Reload()
        Else
            If Not Me.Database.LastQuery.Successful Then
                Throw New Exception(Me.Database.LastQuery.ErrorMsg)
            Else
                Throw New Exception("Opportunity with id #" & Id & " was not found.")
            End If
        End If
    End Sub

    Public Sub Save()
        Dim Sql As String = ""
        ' Form query
        If Me.Id = 0 Then
            ' New
            Sql &= "INSERT INTO opportunity (amount, bill_to_no, contact_id, date_actual_close, date_created,"
            Sql &= " date_expected_close, date_last_updated, referral, last_updated_by, lead_source_id,"
            Sql &= "name, notes, office, probability, reason_lost, salesperson, ship_to_no, stage_id, type)"
            Sql &= " VALUES (@amount, @bill_to_no, @contact_id, @date_actual_close,"
            Sql &= " @date_expected_close, @date_last_updated, @referral,"
            Sql &= " @last_updated_by, @lead_source_id, "
            Sql &= " @name, @notes, @office, @probability, @reason_lost,"
            Sql &= " @salesperson, @ship_to_no, @stage_id, @type)"

        Else
            ' Update
            Sql &= "UPDATE opportunity SET"
            Sql &= "amount=@amount, bill_to_no=@bill_to_no, contact_id=@contact_id, date_actual_close=@date_actual_close,"
            Sql &= " date_expected_close=@date_expected_close, date_last_updated=@date_last_updated, referral=@referral,"
            Sql &= " last_updated_by=@last_updated_by, lead_source_id=@lead_source_id, "
            Sql &= " name=@name, notes=@notes, office=@office, probability=@probability, reason_lost=@reason_lost,"
            Sql &= " salesperson=@salesperson, ship_to_no=@ship_to_no, stage_id=@stage_id, type=@type"
            Sql &= " WHERE id=" & Me._Id
        End If
        ' Put in paramenters
        Sql = Sql.Replace("@amount", Me.Database.Escape(Me.Amount))
        Sql = Sql.Replace("@bill_to_no", Me.Database.Escape(Me.BillToNo))
        Sql = Sql.Replace("@contact_id", Me.Database.Escape(Me.ContactId))
        Sql = Sql.Replace("@date_actual_close", Me.Database.Escape(Me.Database.Escape(IIf(Me.DateActualClose = Nothing, DBNull.Value, Me.DateActualClose))))
        Sql = Sql.Replace("@date_created", Me.Database.Escape(Me.DateCreated))
        Sql = Sql.Replace("@date_expected_close", Me.Database.Escape(IIf(Me.DateExpectedClose = Nothing, DBNull.Value, Me.DateExpectedClose)))
        Sql = Sql.Replace("@date_last_updated", Me.Database.Escape(Me.DateLastUpdated))
        Sql = Sql.Replace("@referral", Me.Database.Escape(Me.EmployeeReferral))
        Sql = Sql.Replace("@last_updated_by", Me.Database.Escape(Me.LastUpdatedBy))
        Sql = Sql.Replace("@lead_source_id", Me.Database.Escape(Me.LeadSourceId))
        Sql = Sql.Replace("@name", Me.Database.Escape(Me.Name))
        Sql = Sql.Replace("@notes", Me.Database.Escape(Me.Notes))
        Sql = Sql.Replace("@office", Me.Database.Escape(Me.Office))
        Sql = Sql.Replace("@probability", Me.Database.Escape(Me.Probability))
        Sql = Sql.Replace("@reason_lost", Me.Database.Escape(Me.ReasonLost))
        Sql = Sql.Replace("@salesperson", Me.Database.Escape(Me.Salesperson))
        Sql = Sql.Replace("@ship_to_no", Me.Database.Escape(Me.ShipToNo))
        Sql = Sql.Replace("@stage_id", Me.Database.Escape(Me.StageId))
        Sql = Sql.Replace("@type", Me.Database.Escape(Me.Type))
        ' Process query
        If Me.Id = Nothing Then
            Me.Database.InsertAndReturnId(Sql)
        Else
            Me.Database.Execute(Sql)
        End If
        RaiseEvent Saved(Me)
        ' Process query success or failure
        If Not Me.Database.LastQuery.Successful Then
            Throw New Exception(Me.Database.LastQuery.ErrorMsg)
        Else
            If Me.Id = Nothing Then
                Me._Id = Me.Database.LastQuery.InsertId
            End If
            Me.Open(Me._Id)
        End If
    End Sub

End Class
