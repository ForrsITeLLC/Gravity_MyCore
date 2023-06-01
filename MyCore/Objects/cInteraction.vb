Imports MyCore.Data

Public Class cInteraction

    Dim _Id As Integer = Nothing
    Dim Database As MyCore.Data.EasySql

    Public EntryType As CommunicationType = CommunicationType.PhoneCall
    Public Subject As String = ""
    Public Body As String = ""
    Public CreatedBy As String = ""
    Public DateCreated As Date = Today
    Public Employee As String = ""
    Public DateOfInteraction As Date = Today
    Public ReferenceType As ReferenceTypes = ReferenceTypes.Sales
    Public OrderNo As String = Nothing
    Public Initiator As InitiatorType = InitiatorType.Us
    Public LastUpdatedBy As String = ""
    Public DateLastUpdated As Date = Today
    Public ContactID As Integer = 0
    Public ContactName As String = ""
    Public CustomerNo As String = ""

    Public Enum InitiatorType
        Us = 1
        Them = 2
    End Enum

    Public Enum ReferenceTypes
        Sales = 1
        Service = 2
        Invoice = 3
        Rental = 4
        PO = 5
        Opportunity = 6
    End Enum

    Public Enum CommunicationType
        NotApplicable = 0
        PhoneCall = 1
        Email = 2
        Fax = 3
        InPerson = 4
        SnailMail = 5
        SendGift = 6
        WebSite = 7
        Other = 8
        SocialNetworkingSite = 9
        InstantMesasge = 10
        TextMessage = 11
    End Enum

    Public Event Reload()
    Public Event Saved(ByVal Interaction As cInteraction)

    Public ReadOnly Property Id() As Integer
        Get
            Return Me._Id
        End Get
    End Property

    Public ReadOnly Property EntryTypes() As DataTable
        Get
            Return Me.Database.GetAll("SELECT id, name FROM journal_entry_type WHERE archived = 0")
        End Get
    End Property

    Public Sub New(ByRef db As MyCore.Data.EasySql)
        Me.Database = db
    End Sub

    Public Sub Open(ByVal Id As Integer)
        Dim Sql As String = ""
        Sql &= " SELECT i.contact_name, i.customer_no, i.id, i.entry_type_id, i.subject, i.memo, i.contact_id,"
        Sql &= " i.created_by, i.created_date, i.touch_date, i.touch_by, i.department, i.initiator,"
        Sql &= " i.date_last_updated, i.last_updated_by, i.ref_no, c.cnt_last + ', ' + c.cnt_first AS display_name"
        Sql &= " FROM journal i LEFT OUTER JOIN CONTACTS c ON i.contact_id=c.cnt_id"
        Sql &= " WHERE id=" & Id
        Dim Row As DataRow = Me.Database.GetRow(Sql)
        If Me.Database.LastQuery.Successful And Me.Database.LastQuery.RowsReturned = 1 Then
            Me._Id = Id
            Me.ContactID = IIf(Row.Item("contact_id") Is DBNull.Value, 0, Row.Item("contact_id"))
            Me.ContactName = Row.Item("contact_name")
            Me.CustomerNo = Row.Item("customer_no")
            Me.EntryType = Row.Item("entry_type_id")
            Me.Subject = Row.Item("subject")
            Me.Body = Row.Item("memo")
            Me.CreatedBy = Row.Item("created_by")
            Me.DateCreated = Row.Item("created_date")
            Me.Employee = Row.Item("touch_by")
            Me.DateOfInteraction = Row.Item("touch_date")
            Me.ReferenceType = Row.Item("department")
            Me.OrderNo = IIf(Row.Item("ref_no") Is DBNull.Value, "", Row.Item("ref_no"))
            Me.Initiator = Row.Item("initiator")
            Me.DateLastUpdated = Row.Item("date_last_updated")
            Me.LastUpdatedBy = Row.Item("last_updated_by")
            RaiseEvent Reload()
        ElseIf Not Me.Database.LastQuery.Successful Then
            Throw New Exception(Me.Database.LastQuery.ErrorMsg)
        Else
            Throw New Exception("Interaction id#" & Id & " not found.")
        End If
    End Sub

    Public Sub Save()
        Dim Sql As String = ""
        ' Form query
        If Me.Id = Nothing Then
            ' New
            Sql &= "INSERT INTO journal (contact_name, contact_id, customer_no, entry_type_id, subject, memo,"
            Sql &= " created_by, created_date, touch_date, touch_by, department, initiator,"
            Sql &= " date_last_updated, last_updated_by, ref_no)"
            Sql &= " VALUES (@contact_name, @contact_id, @customer_no, @entry_type_id, @subject, @memo,"
            Sql &= " @created_by, @date_created, @touch_date, @touch_by, @department, @initiator,"
            Sql &= " @date_last_updated, @last_updated_by, @ref_no)"

        Else
            ' Update
            Sql &= "UPDATE journal SET"
            Sql &= " contact_name=@contact_name, contact_id=@contact_id, customer_no=@customer_no, entry_type_id=@entry_type_id,"
            Sql &= " subject=@subject, memo=@memo, touch_date=@touch_date, touch_by=@touch_by,"
            Sql &= " department=@department, initiator=@initiator, date_last_updated=@date_last_updated,"
            Sql &= " last_updated_by=@last_updated_by, ref_no=@ref_no"
            Sql &= " WHERE id=" & Me._Id
        End If
        ' Put in paramenters
        Sql = Sql.Replace("@contact_name", Me.Database.Escape(Me.ContactName))
        Sql = Sql.Replace("@contact_id", Me.Database.Escape(Me.ContactID))
        Sql = Sql.Replace("@customer_no", Me.Database.Escape(Me.CustomerNo))
        Sql = Sql.Replace("@entry_type_id", Me.Database.Escape(CInt(Me.EntryType)))
        Sql = Sql.Replace("@subject", Me.Database.Escape(Me.Subject))
        Sql = Sql.Replace("@memo", Me.Database.Escape(Me.Body))
        Sql = Sql.Replace("@created_by", Me.Database.Escape(Me.CreatedBy))
        Sql = Sql.Replace("@date_created", Me.Database.Escape(Me.DateCreated))
        Sql = Sql.Replace("@touch_date", Me.Database.Escape(Me.DateOfInteraction))
        Sql = Sql.Replace("@touch_by", Me.Database.Escape(Me.Employee))
        Sql = Sql.Replace("@department", Me.Database.Escape(CInt(Me.ReferenceType)))
        Sql = Sql.Replace("@initiator", Me.Database.Escape(CInt(Me.Initiator)))
        Sql = Sql.Replace("@date_last_updated", Me.Database.Escape(Me.DateLastUpdated))
        Sql = Sql.Replace("@last_updated_by", Me.Database.Escape(Me.LastUpdatedBy))
        Sql = Sql.Replace("@ref_no", Me.Database.Escape(Me.OrderNo))
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
