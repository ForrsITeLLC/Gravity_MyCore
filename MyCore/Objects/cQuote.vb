Public Class cQuote

    Dim _Id As Integer = 0

    Dim _ShipToName As String = ""
    Dim _BillToName As String = ""

    Public Name As String = ""
    Public Attention As String = ""
    Public OpportunityId As Integer = 0
    Public RevisionOf As Integer = 0
    Public DateSent As Date = Today
    Public DateExpires As Date = Nothing
    Public TaxGroupId As Integer = 0
    Public TermsId As Integer = 0
    Public Fob As String = ""
    Public ShipViaId As Integer = 0
    Public Freight As Double = 0
    Public DateCreated As Date = Today
    Public DateLastUpdated As Date = Today
    Public LastUpdatedBy As String = ""
    Public CreatedBy As String = ""
    Public RateTableId As Integer = 0
    Public Voided As Boolean = False

    Public LineItems As DataTable

    Public Database As MyCore.Data.EasySql

    Public Event Reload()


    Public ReadOnly Property Id() As Integer
        Get
            Return Me._Id
        End Get
    End Property

    Public ReadOnly Property ShipToName() As String
        Get
            Return Me._ShipToName
        End Get
    End Property

    Public ReadOnly Property BillToName() As String
        Get
            Return Me._BillToName
        End Get
    End Property

    Public ReadOnly Property TaxStatuses() As DataTable
        Get
            Return Me.Database.GetAll("SELECT id, code, description, taxable FROM tax_status ORDER BY code")
        End Get
    End Property

    Public ReadOnly Property ShipVias() As DataTable
        Get
            Return Me.Database.GetAll("SELECT id, name, sort FROM ship_via ORDER BY sort")
        End Get
    End Property

    Public ReadOnly Property TaxGroups() As DataTable
        Get
            Return Me.Database.GetAll("SELECT id, name FROM tax_code ORDER BY name")
        End Get
    End Property

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
    End Sub

    Public Sub Open(ByVal Id As Integer)
        Dim Sql As String = ""
        Sql &= "SELECT quote.*, b.cst_name AS bill_to_name, s.cst_name = ship_to_name FROM quote"
        Sql &= " LEFT JOIN opportunity o ON quote.opportunity_id=opportunity.id"
        Sql &= " LEFT JOIN ADDRESS b ON quote.bill_to_no=b.cst_no"
        Sql &= " LEFT JOIN ADDRESS s ON quote.ship_to_no=s.cst_no"
        Sql &= " WHERE id=" & Id
        Dim Quote As DataRow = Me.Database.GetRow(Sql)
        If Me.Database.LastQuery.Successful And Me.Database.LastQuery.RowsReturned = 1 Then
            Me.Name = Quote.Item("quote_name")
            Me._BillToName = Quote.Item("bill_to_name")
            Me._ShipToName = Quote.Item("ship_to_name")
            Me.Attention = Quote.Item("sent_to_name")
            Me.OpportunityId = Quote.Item("opportunity_id")
            Me.RevisionOf = Quote.Item("revision_of")
            Me.ShipViaId = Quote.Item("ship_via_id")
            Me.DateSent = Quote.Item("date_sent")
            Me.DateExpires = Quote.Item("date_expires")
            Me.Freight = Quote.Item("freight")
            Me.TaxGroupId = Quote.Item("tax_code_id")
            Me.RateTableId = Quote.Item("rate_table_id")
            Me.DateCreated = Quote.Item("date_created")
            Me.CreatedBy = Quote.Item("created_by")
            Me.DateLastUpdated = Quote.Item("date_last_updated")
            Me.LastUpdatedBy = Quote.Item("last_updated_by")
            Me.TermsId = Quote.Item("terms_id")
            Me.Fob = Quote.Item("fob")
            Me.Voided = Quote.Item("voided")
            RaiseEvent Reload()
        Else
            If Not Me.Database.LastQuery.Successful Then
                Throw New Exception(Me.Database.LastQuery.ErrorMsg)
            Else
                Throw New Exception("Quote with id #" & Id & " was not found.")
            End If
        End If
    End Sub

    Public Sub Save()

    End Sub

End Class
