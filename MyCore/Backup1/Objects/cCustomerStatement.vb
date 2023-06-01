Public Class cCustomerStatement

    Dim Database As MyCore.Data.EasySql
    Dim Settings As cSettings

    Public LateMessage As String = "YOUR ACCOUNT IS PAST DUE! Please remit. Please call immediately if this statement is in error."
    Public CurrentMessage As String = "Thank you for your business and for keeping invoices current."
    Public _TemplateId As Integer = 0
    Public PageHeightInPixels As Integer = 912

    Dim Html As String = ""

    Public Property TemplateId() As Integer
        Get
            Return Me._TemplateId
        End Get
        Set(ByVal value As Integer)
            Me._TemplateId = value
            Me.Html = Me.Database.GetOne("SELECT html FROM template WHERE id=" & TemplateId)
        End Set
    End Property

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
        Me.Settings = New cSettings(Me.Database)
        Me.CurrentMessage = Me.Settings.GetValue("Statement Current Message", "Thank you for your business and for keeping invoices current.")
        Me.LateMessage = Me.Settings.GetValue("Statement Late Message", "YOUR ACCOUNT IS PAST DUE! Please remit. Please call immediately if this statement is in error.")
        Me.TemplateId = Me.Settings.GetValue("Template Statement", 0)
        Me.PageHeightInPixels = Me.Settings.GetValue("Page Height in Pixels")
    End Sub

    Public Function GetMultipleCustomerStatements(ByVal TemplateId As Integer, ByVal Customers As String()) As GravityDocument.gDocument
        Dim Doc As New GravityDocument.gDocument(PageHeightInPixels)
        For Each no As String In Customers
            Dim Customer As New cCompany(Me.Database)
            Customer.Open(no)
            Doc.AddPage(Me.ToGravityDocument(Customer).Pages(1))
        Next
        Return Doc
    End Function

    Public Function ToGravityDocument(ByRef Customer As cCompany) As MyCore.GravityDocument.gDocument
        System.Windows.Forms.Application.DoEvents()
        ' Get company
        Dim CustomerName As String = ""
        Dim CustomerAddress As String = ""
        Dim CustomerCity As String = ""
        Dim CustomerState As String = ""
        Dim CustomerZip As String = ""
        ' If bill to has a billing address use it
        If Customer.BillingAddress1.Length > 0 Then
            CustomerName = Customer.BillingName
            CustomerAddress = Customer.BillingAddress1
            If Customer.BillingAddress2.Length > 0 Then
                CustomerAddress &= ControlChars.CrLf & Customer.BillingAddress2
            End If
            CustomerCity = Customer.BillingCity
            CustomerState = Customer.BillingState
            CustomerZip = Customer.BillingZip
        Else
            CustomerName = Customer.Name
            CustomerAddress = Customer.Address1
            If Customer.Address2.Length > 0 Then
                CustomerAddress &= ControlChars.CrLf & Customer.Address2
            End If
            CustomerCity = Customer.City
            CustomerState = Customer.State
            CustomerZip = Customer.Zip
        End If
        ' Create Template
        Dim Doc As New GravityDocument.gDocument(Me.PageHeightInPixels)
        Dim Page As GravityDocument.gPage = Doc.AddPageFromXml(Me.Html)
        ' Populate customer information
        Page.AddVariable("%customer_no%", Customer.CustomerNo)
        Page.AddVariable("%customer_name%", CustomerName)
        Page.AddVariable("%address%", CustomerAddress)
        Page.AddVariable("%city%", CustomerCity)
        Page.AddVariable("%state%", CustomerState)
        Page.AddVariable("%zip%", CustomerZip)
        Page.AddVariable("%phone%", Customer.Phone)
        Page.AddVariable("%fax%", Customer.Fax)
        Page.AddVariable("%email%", Customer.APEmailAddress)
        Page.AddVariable("%date%", Today.ToString("MM/dd/yy"))
        ' Get Statement
        Dim Statement As DataTable() = Customer.StatementTable
        ' Populate breakdown of days
        Page.AddVariable("%current%", Format(CType(Statement(1).Rows(0).Item("current"), Double), "c"))
        Page.AddVariable("%31to60%", Format(CType(Statement(1).Rows(0).Item("31to60"), Double), "c"))
        Page.AddVariable("%61to90%", Format(CType(Statement(1).Rows(0).Item("61to90"), Double), "c"))
        Page.AddVariable("%91plus%", Format(CType(Statement(1).Rows(0).Item("91plus"), Double), "c"))
        ' Show payment message
        If Statement(1).Rows(0).Item("31to60") > 0 Or Statement(1).Rows(0).Item("61to90") > 0 Or Statement(1).Rows(0).Item("91plus") > 0 Then
            Page.AddVariable("%payment_message%", Me.LateMessage)
        Else
            Page.AddVariable("%payment_message%", Me.CurrentMessage)
        End If
        ' Populate table
        Page.GetTableBySource("line_items").Table.Data = Statement(0)
        ' Load it
        Return Doc
    End Function

End Class
