Imports MyCore.Data

Public Class cInvoice

    Dim _Id As Integer = 0
    Dim _Freight As Double = 0
    Public InvoiceNo As String = ""
    Public BillTo As String = ""
    Public ShipTo As String = ""
    Public Office As String = 0
    Public DateInvoiced As Date = Nothing
    Public TermsId As Integer = 0
    Public CustomerPo As String = ""
    Public OrderNo As String = ""
    Public OrderType As TypeOfOrder = TypeOfOrder.Sales
    Public DateDue As Date = Nothing
    Public Salesman As String = ""
    Public Printed As Boolean = False
    Public Exported As Boolean = False
    Public TaxCodeId As Integer = 0
    Public SalesTax As Double = 0
    Public LateChargePercent As Double = 2
    Public Memo As String = ""
    Public DateLastUpdated As DateTime = Nothing
    Public LastUpdatedBy As String = ""
    Public Paid As Boolean = False
    Public DateCreated As DateTime = Nothing
    Public CreatedBy As String = ""
    Public LineItems As DataTable
    Public GeneralLedger As DataTable
    Public Commissions As DataTable
    Public WarrantyCredit As Boolean = False
    Public ProjectId As Integer = 0
    Public IsCreditMemo As Boolean = False
    Public AcctRef As String = Nothing
    Public Voided As Boolean = False

    Public TermsRef As String = Nothing
    Public SalesmanRef As String = Nothing
    Public TaxCodeRef As String = Nothing
    Public BillToRef As String = Nothing
    Public ShipToRef As String = Nothing
    Public TaxRef As String = Nothing

    Dim _Office As DataTable

    Dim Database As MyCore.Data.EasySql

    Public Event Reload()
    Public Event Saved(ByVal Invoice As cInvoice)

    Public Enum TypeOfOrder As Integer
        Other = 0
        Service = 1
        Sales = 2
        Rental = 3
    End Enum

    Public ReadOnly Property Id() As Integer
        Get
            Return Me._Id
        End Get
    End Property

    Public ReadOnly Property Offices() As DataTable
        Get
            Return Me._Office
        End Get
    End Property

    Public ReadOnly Property OfficeRef() As String
        Get
            Dim Ref As String
            Select Case Me.OrderType
                Case TypeOfOrder.Sales
                    Ref = Me.Database.GetOne("SELECT sales_ref FROM office WHERE number=" & Me.Database.Escape(Me.Office))
                Case TypeOfOrder.Service
                    Ref = Me.Database.GetOne("SELECT service_ref FROM office WHERE number=" & Me.Database.Escape(Me.Office))
                Case TypeOfOrder.Rental
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

    Public ReadOnly Property ContactPerson() As String
        Get
            If Me.OrderNo > 0 Then
                If Me.OrderType = TypeOfOrder.Service Then
                    Dim Sql As String = "SELECT contact_name FROM service_order WHERE id=" & Me.OrderNo
                    Return Me.Database.GetOne(Sql)
                ElseIf Me.OrderType = TypeOfOrder.Sales Then
                    Dim Sql As String = "SELECT contact FROM sales_order WHERE id=" & Me.OrderNo
                    Return Me.Database.GetOne(Sql)
                Else
                    Dim Sql As String = "SELECT contact FROM rental_order WHERE id=" & Me.OrderNo
                    Return Me.Database.GetOne(Sql)
                End If
            Else
                Return ""
            End If
        End Get
    End Property

    Public Property Freight() As Double
        Get
            Return Math.Round(Me._Freight, 2)
        End Get
        Set(ByVal value As Double)
            Me._Freight = Math.Round(value, 2)
        End Set
    End Property

    Public ReadOnly Property AmountPaid() As Double
        Get
            Dim Paid As Double = Me.Database.GetOne("SELECT ISNULL(SUM(pgl_amt), 0) FROM PYMNTGLD WHERE pgl_ino=" & Me.Database.Escape(Me.InvoiceNo))
            If Me.IsCreditMemo Then
                Return Math.Round(Paid * -1, 2)
            Else
                Return Math.Round(Paid, 2)
            End If
        End Get
    End Property

    Public ReadOnly Property DateLastPayment() As Date
        Get
            Dim Row As DataRow = Me.Database.GetRow("SELECT MAX(pmt_date) As last_payment FROM PAYMENTS LEFT JOIN PYMNTGLD ON pmt_recno=pgl_recno WHERE pgl_ino=" & Me.Database.Escape(Me.InvoiceNo))
            If Row.Item("last_payment") Is DBNull.Value Then
                Return Nothing
            Else
                Return Row.Item("last_payment")
            End If
        End Get
    End Property

    Public ReadOnly Property Subtotal() As Double
        Get
            Dim total As Double = 0
            For Each r As DataRow In Me.LineItems.Rows
                Try
                    total += r.Item("quantity") * r.Item("unit_price")
                Catch
                End Try
            Next
            Return Math.Round(total, 2)
        End Get
    End Property

    Public ReadOnly Property Total() As Double
        Get
            Return Math.Round(Me.Subtotal + Me.Freight + Me.SalesTax, 2)
        End Get
    End Property

    Public ReadOnly Property TotalPlusLateFee() As Double
        Get
            Return Me.Total * Math.Round((1 + (Me.LateChargePercent / 100)), 2)
        End Get
    End Property

    Public ReadOnly Property Balance() As Double
        Get
            Dim InvoiceTotal As Double = Me.Total
            Dim PaidTotal As Double = Me.AmountPaid
            Return InvoiceTotal - PaidTotal
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
            Sql &= " WHERE ref_no=" & Me.Database.Escape(Me.InvoiceNo)
            Sql &= " AND department=" & CInt(cInteraction.ReferenceTypes.Invoice)
            Return Me.Database.GetAll(Sql)
        End Get
    End Property

    Public ReadOnly Property Payments() As DataTable
        Get
            Dim Sql As String = "SELECT DISTINCT pmt_recno AS record_no, pmt_date AS payment_date, pmt_chck AS check_no,"
            Sql &= " amount = SUM(pgl_amt)"
            Sql &= " FROM PYMNTGLD"
            Sql &= " INNER JOIN PAYMENTS ON pgl_recno=pmt_recno"
            Sql &= " WHERE pgl_ino=" & Me.Database.Escape(Me.InvoiceNo)
            Sql &= " GROUP BY pmt_recno, pmt_date, pmt_chck"
            Return Me.Database.GetAll(Sql)
        End Get
    End Property

    Public Sub New(ByRef db As MyCore.Data.EasySql)
        Me.Database = db
        Me.PopulateOffice()
        ' Make General Ledger Table
        Dim Table As New DataTable
        Table.Columns.Add("id")
        Table.Columns.Add("account_no")
        Table.Columns.Add("employee")
        Table.Columns.Add("amount")
        Table.Columns.Add("exported")
        Table.Columns.Add("date_created")
        Table.Columns.Add("date_last_updated")
        Table.Columns.Add("created_by")
        Table.Columns.Add("last_updated_by")
        Me.GeneralLedger = Table
        ' Make Commissions Table
        Dim Table2 As New DataTable
        Table2.Columns.Add("com_id")
        Table2.Columns.Add("com_no")
        Table2.Columns.Add("com_amt")
        Table2.Columns.Add("hold")
        Table2.Columns.Add("com_user")
        Table2.Columns.Add("com_chngt")
        Me.Commissions = Table2
        ' Make Line Items Table
        Dim Table3 As New DataTable
        Table3.Columns.Add("id")
        Table3.Columns.Add("quantity")
        Table3.Columns.Add("part_no")
        Table3.Columns.Add("serial_no")
        Table3.Columns.Add("description")
        Table3.Columns.Add("unit_price")
        Table3.Columns.Add("unit_cost")
        Table3.Columns.Add("tax_status_id")
        Table3.Columns.Add("taxable")
        Table3.Columns.Add("user")
        Table3.Columns.Add("updated")
        Table3.Columns.Add("ext_price")
        Me.LineItems = Table3
    End Sub

    Private Sub PopulateOffice()
        Me._Office = Me.Database.GetAll("SELECT id, number, name, sort FROM office ORDER BY sort")
    End Sub

    Public Sub Open(ByVal InvoiceNumber As String)
        Dim Row As DataRow
        Dim Sql As String = ""
        Sql &= " SELECT inv.*, bt.acct_ref AS bill_to_acct_ref, tc.acct_ref AS tax_ref,"
        Sql &= " st.acct_ref As ship_to_acct_ref, "
        Sql &= " terms.acct_ref AS terms_acct_ref, emp.acct_ref AS employee_acct_ref"
        Sql &= " FROM INVOICE inv"
        Sql &= " LEFT OUTER JOIN ADDRESS bt ON inv.inv_no=bt.cst_no"
        Sql &= " LEFT OUTER JOIN ADDRESS st ON inv.ship_to_no=st.cst_no"
        Sql &= " LEFT OUTER JOIN pay_status terms ON inv.terms_id=terms.id"
        Sql &= " LEFT OUTER JOIN employee emp ON inv.inv_slsmn=emp.windows_user"
        Sql &= " LEFT OUTER JOIN tax_code tc ON inv.tax_code_id=tc.id"
        Sql &= " WHERE inv.inv_ino=@invoice_no"
        Sql = Sql.Replace("@invoice_no", Me.Database.Escape(InvoiceNumber))
        Row = Me.Database.GetRow(Sql)
        If Me.Database.LastQuery.RowsReturned = 1 Then
            Me.BillToRef = Me.IsNull(Row.Item("bill_to_acct_ref"), Nothing)
            Me.ShipToRef = Me.IsNull(Row.Item("ship_to_acct_ref"), Nothing)
            Me.TermsRef = Me.IsNull(Row.Item("terms_acct_ref"), Nothing)
            Me.SalesmanRef = Me.IsNull(Row.Item("employee_acct_ref"), Nothing)
            ' Set values
            Me._Id = Row.Item("inv_id")
            Me.InvoiceNo = Row.Item("inv_ino")
            Me.IsCreditMemo = Row.Item("refund")
            Me.BillTo = Me.IsNull(Row.Item("inv_no"), "")
            Me.ShipTo = Me.IsNull(Row.Item("ship_to_no"), "")
            Me.Office = Me.IsNull(Row.Item("inv_off"), 0)
            Me.DateInvoiced = Me.IsNull(Row.Item("inv_date"), Nothing)
            Me.DateDue = Me.IsNull(Row.Item("date_due"), Nothing)
            Me.TermsId = Me.IsNull(Row.Item("terms_id"), 0)
            Me.CustomerPo = Me.IsNull(Row.Item("inv_yord"), "")
            Me.OrderNo = Me.IsNull(Row.Item("inv_ourd"), "")
            Me.OrderType = Row.Item("order_type")
            Me.Salesman = Me.IsNull(Row.Item("inv_slsmn"), "")
            Me.Printed = IIf(Row.Item("inv_prnt") = "Y", False, True)
            If Me.IsCreditMemo Then
                Me.SalesTax = Row.Item("sales_tax") * -1
                Me.Freight = Me.IsNull(Row.Item("inv_shch"), 0) * -1
            Else
                Me.SalesTax = Row.Item("sales_tax")
                Me.Freight = Me.IsNull(Row.Item("inv_shch"), 0)
            End If
            Me.LateChargePercent = Me.IsNull(Row.Item("inv_lchp"), 2)
            Me.Memo = Me.IsNull(Row.Item("inv_lchc"), "")
            Me.DateLastUpdated = Me.IsNull(Row.Item("inv_chngt"), Nothing)
            Me.LastUpdatedBy = Me.IsNull(Row.Item("inv_user"), "")
            Me.Exported = Row.Item("inv_exported")
            Me.Paid = Row.Item("inv_paid")
            Me.WarrantyCredit = Row.Item("warranty_billing")
            Me.OrderType = Row.Item("order_type")
            Me.TaxCodeId = Row.Item("tax_code_id")
            Me.TaxCodeRef = Me.IsNull(Row.Item("tax_ref"), Nothing)
            Me.Voided = Row.Item("voided")
            Me.AcctRef = Me.IsNull(Row.Item("acct_ref"), Nothing)
            Me.LineItems = Me.GetLineItems
            Sql = "SELECT id, account_no, employee,"
            If Me.IsCreditMemo Then
                Sql &= " (amount*-1) AS amount, "
            Else
                Sql &= " amount, "
            End If
            Sql &= " exported, date_created, date_last_updated, created_by, last_updated_by"
            Sql &= " FROM invoice_gl"
            Sql &= " WHERE invoice_no=" & Me.Database.Escape(InvoiceNumber)
            Me.GeneralLedger = Me.Database.GetAll(Sql)
            Sql = "SELECT com_id, com_no, "
            If Me.IsCreditMemo Then
                Sql &= " (com_amt*-1) AS com_amt, (hold*-1) AS hold,"
            Else
                Sql &= " com_amt, hold, "
            End If
            Sql &= " com_chngt, com_user"
            Sql &= " FROM COMMIS "
            Sql &= " WHERE com_inv=" & Me.Database.Escape(InvoiceNumber)
            Me.Commissions = Me.Database.GetAll(Sql)
            RaiseEvent Reload()
        ElseIf Not Me.Database.LastQuery.Successful Then
            Throw New Exception(Me.Database.LastQuery.ErrorMsg)
        Else
            Throw New Exception("Invoice not found.")
        End If
    End Sub

    Private Function GetLineItems() As DataTable
        Dim Sql As String = ""
        Sql &= " SELECT ivi_id AS [id], "
        Sql &= " ivi_qua AS [quantity], "
        Sql &= " ivi_ser AS [serial_no], "
        Sql &= " ivi_desc AS [description], "
        If Me.IsCreditMemo Then
            Sql &= " ISNULL(ivi_cost, 0)*-1 AS [unit_price], "
            Sql &= " ISNULL(ivi_ocost, 0)*-1 AS [unit_cost], "
        Else
            Sql &= " ISNULL(ivi_cost, 0) AS [unit_price], "
            Sql &= " ISNULL(ivi_ocost, 0) AS [unit_cost], "
        End If
        Sql &= " ivi_user AS [user], "
        Sql &= " ivi_chngt AS [updated],"
        If Me.IsCreditMemo Then
            Sql &= " ISNULL([ivi_qua] * [ivi_cost], 0)*-1 AS [ext_price], "
        Else
            Sql &= " ISNULL([ivi_qua] * [ivi_cost], 0)  AS [ext_price], "
        End If
        Sql &= " ivi_part AS part_no, item_master.acct_ref AS part_no_ref,"
        Sql &= " item_master.item_type_id,"
        Sql &= " INVITEMS.tax_status_id, tax_status.acct_ref AS tax_ref, taxable"
        Sql &= " FROM INVITEMS"
        Sql &= " LEFT OUTER JOIN tax_status ON tax_status.id=INVITEMS.tax_status_id"
        Sql &= " LEFT OUTER JOIN item_master ON item_master.part_no=INVITEMS.ivi_part"
        Sql &= " WHERE ivi_ino=" & Me.Database.Escape(Me.InvoiceNo)
        Dim Li As DataTable = Me.Database.GetAll(Sql)
        If Me.Database.LastQuery.Successful Then
            Return Li
        Else
            Dim er As String = Me.Database.LastQuery.ErrorMsg
            Throw New Exception("Did not retrieve line items. " & er)
        End If
    End Function

    Private Function IsNull(ByVal Value As Object, Optional ByVal ReturnValue As Object = "") As Object
        If Value Is DBNull.Value Then
            Return ReturnValue
        Else
            Return Value
        End If
    End Function

    Public Sub Void()
        ' Mark it voided
        Me.Voided = True
        ' Save
        Me.Save()
        ' Mark orders billed to this invoice as unbilled
        Me.Database.Execute("UPDATE service_order SET invoice_id=0 WHERE charge_to=0 AND invoice_id=" & Me.Database.Escape(Me.InvoiceNo))
        Me.Database.Execute("UPDATE sales_order SET invoice_no='' WHERE invoice_no=" & Me.Database.Escape(Me.InvoiceNo))
        Me.Database.Execute("UPDATE rental_order SET invoice_no='' WHERE invoice_no=" & Me.Database.Escape(Me.InvoiceNo))
    End Sub

    Public Sub Save()
        If Me.Total < 0 Then
            Throw New Exception("Cannot have an invoice for a negative amount.")
        End If
        Dim Sql As String = ""
        If Me._Id > 0 Then
            'Save Existing Service Order
            Sql = "UPDATE INVOICE "
            Sql &= " SET"
            Sql &= " inv_ino=@invoice_no,"
            Sql &= " inv_no=@bill_to,"
            Sql &= " inv_date=@invoice_date, "
            Sql &= " inv_off=@office, "
            Sql &= " inv_yord=@customer_po, "
            Sql &= " inv_ourd=@order_no,"
            Sql &= " inv_slsmn=@salesman, "
            Sql &= " sales_tax=" & Me.Database.ToCurrency("@sales_tax") & ", "
            Sql &= " tax_code_id=@tax_code_id,"
            Sql &= " inv_shch=" & Me.Database.ToCurrency("@freight") & ", "
            Sql &= " inv_lchp=@late_charge_percent,"
            Sql &= " inv_lchc=@memo,"
            Sql &= " inv_paid=@paid, "
            Sql &= " inv_exported=@exported, "
            Sql &= " order_type=@order_type,"
            Sql &= " inv_prnt=@print,"
            Sql &= " date_due=@date_due,"
            Sql &= " ship_to_no=@ship_to,"
            Sql &= " warranty_billing=@warranty_credit,"
            Sql &= " terms_id=@terms_id,"
            Sql &= " refund=@refund,"
            Sql &= " voided=@voided"
            Sql &= " WHERE inv_id=@invoice_id"
            Sql = Sql.Replace("@invoice_id", Me._Id)
            Sql = Sql.Replace("@invoice_no", Me.Database.Escape(Me.InvoiceNo))
            Sql = Sql.Replace("@invoice_date", Me.Database.Escape(Me.DateInvoiced))
            Sql = Sql.Replace("@date_due", Me.Database.Escape(Me.DateDue))
            Sql = Sql.Replace("@ship_to", Me.Database.Escape(Me.ShipTo))
            Sql = Sql.Replace("@bill_to", Me.Database.Escape(Me.BillTo))
            Sql = Sql.Replace("@office", Me.Database.Escape(Me.Office))
            Sql = Sql.Replace("@terms_id", Me.TermsId)
            Sql = Sql.Replace("@customer_po", Me.Database.Escape(Me.CustomerPo))
            Sql = Sql.Replace("@order_no", Me.Database.Escape(Me.OrderNo))
            Sql = Sql.Replace("@order_type", Me.OrderType)
            Sql = Sql.Replace("@salesman", Me.Database.Escape(Me.Salesman))
            Sql = Sql.Replace("@tax_code_id", Me.TaxCodeId)
            If Me.IsCreditMemo Then
                Sql = Sql.Replace("@sales_tax", Me.SalesTax * -1)
                Sql = Sql.Replace("@freight", Me.Freight * -1)
            Else
                Sql = Sql.Replace("@sales_tax", Me.SalesTax)
                Sql = Sql.Replace("@freight", Me.Freight)
            End If
            Sql = Sql.Replace("@late_charge_percent", Me.LateChargePercent)
            Sql = Sql.Replace("@exported", Me.Database.Escape(Me.Exported))
            Sql = Sql.Replace("@paid", Me.Database.Escape(Me.Paid))
            Sql = Sql.Replace("@print", Me.Database.Escape(IIf(Me.Printed, "N", "Y")))
            Sql = Sql.Replace("@warranty_credit", Me.Database.Escape(Me.WarrantyCredit))
            Sql = Sql.Replace("@memo", Me.Database.Escape(Me.Memo))
            Sql = Sql.Replace("@voided", Me.Database.Escape(Me.Voided))
            Sql = Sql.Replace("@refund", Me.Database.Escape(Me.IsCreditMemo))
            Me.Database.Execute(Sql)
        Else
            ' New Invoice
            Sql = "INSERT INTO INVOICE"
            Sql &= " (inv_date, inv_no, ship_to_no, inv_off, inv_slsmn, inv_yord,"
            Sql &= " inv_paid, inv_lchp, inv_lchc, inv_shch, inv_ino, inv_prnt, inv_exported,"
            Sql &= " tax_code_id, sales_tax,"
            Sql &= " inv_ourd, order_type, date_due, warranty_billing, terms_id, refund, voided)"
            Sql &= " VALUES"
            Sql &= " (@invoice_date, @bill_to, @ship_to, @office, @salesman, @customer_po,"
            Sql &= " @paid, @late_charge_percent, @memo, " & Me.Database.ToCurrency("@freight") & ", "
            Sql &= " @invoice_no, @print, @exported, @tax_code_id, @sales_tax,"
            Sql &= " @order_no, @order_type, @date_due, @warranty_credit, @terms_id, @refund, @voided)"
            Sql = Sql.Replace("@invoice_id", Me._Id)
            Sql = Sql.Replace("@invoice_no", Me.Database.Escape(Me.InvoiceNo))
            Sql = Sql.Replace("@invoice_date", Me.Database.Escape(Me.DateInvoiced))
            Sql = Sql.Replace("@date_due", Me.Database.Escape(Me.DateDue))
            Sql = Sql.Replace("@ship_to", Me.Database.Escape(Me.ShipTo))
            Sql = Sql.Replace("@bill_to", Me.Database.Escape(Me.BillTo))
            Sql = Sql.Replace("@office", Me.Database.Escape(Me.Office))
            Sql = Sql.Replace("@terms_id", Me.Database.Escape(Me.TermsId))
            Sql = Sql.Replace("@customer_po", Me.Database.Escape(Me.CustomerPo))
            Sql = Sql.Replace("@order_no", Me.Database.Escape(Me.OrderNo))
            Sql = Sql.Replace("@order_type", CInt(Me.OrderType))
            Sql = Sql.Replace("@salesman", Me.Database.Escape(Me.Salesman))
            Sql = Sql.Replace("@tax_code_id", Me.Database.Escape(Me.TaxCodeId))
            If Me.IsCreditMemo Then
                Sql = Sql.Replace("@sales_tax", Me.Database.Escape(Me.SalesTax * -1))
                Sql = Sql.Replace("@freight", Me.Database.Escape(Me.Freight * -1))
            Else
                Sql = Sql.Replace("@sales_tax", Me.Database.Escape(Me.SalesTax))
                Sql = Sql.Replace("@freight", Me.Database.Escape(Me.Freight))
            End If
            Sql = Sql.Replace("@late_charge_percent", Me.Database.Escape(Me.LateChargePercent))
            Sql = Sql.Replace("@exported", Me.Database.Escape(Me.Exported))
            Sql = Sql.Replace("@paid", Me.Database.Escape(Me.Paid))
            Sql = Sql.Replace("@print", Me.Database.Escape(IIf(Me.Printed, "N", "Y")))
            Sql = Sql.Replace("@warranty_credit", Me.Database.Escape(Me.WarrantyCredit))
            Sql = Sql.Replace("@memo", Me.Database.Escape(Me.Memo))
            Sql = Sql.Replace("@refund", Me.Database.Escape(Me.IsCreditMemo))
            Sql = Sql.Replace("@voided", Me.Database.Escape(Me.Voided))
            Me.Database.InsertAndReturnId(Sql)
            If Not Me.Database.LastQuery.Successful Then
                Dim er As String = Me.Database.LastQuery.ErrorMsg
            End If
        End If
        If Me.Database.LastQuery.Successful Then
            If Me._Id = 0 Then
                Me._Id = Me.Database.LastQuery.InsertId
                Me.IncrementNextNumber()
            End If
            ' Save invoice items
            Dim r As DataRow
            For Each r In Me.LineItems.Rows
                Dim PartNo As String = IIf(r.Item("part_no") Is DBNull.Value, "", r.Item("part_no"))
                If r.Item("quantity") Is DBNull.Value Then
                    r.Item("quantity") = 1
                End If
                If r.RowState = DataRowState.Added Then
                    If r.Item("quantity") > 0 Then
                        Sql = "INSERT INTO INVITEMS (ivi_ino, ivi_qua, ivi_desc, ivi_ser, ivi_part, ivi_cost, ivi_ocost, tax_status_id)"
                        Sql &= " VALUES (@invoice_no, @quantity, @description, @serial_no, @part_no, " & Me.Database.ToCurrency("@unit_price") & ", " & Me.Database.ToCurrency("@unit_cost") & ", @tax_status_id)"
                    Else
                        Sql = ""
                    End If
                ElseIf r.RowState = DataRowState.Modified Then
                    If r.Item("quantity") > 0 Then
                        Sql = "UPDATE INVITEMS SET"
                        Sql &= " ivi_qua=@quantity,"
                        Sql &= " ivi_desc=@description,"
                        Sql &= " ivi_part=@part_no,"
                        Sql &= " ivi_ser=@serial_no,"
                        Sql &= " ivi_cost=" & Me.Database.ToCurrency("@unit_price") & ","
                        Sql &= " ivi_ocost=" & Me.Database.ToCurrency("@unit_cost") & ","
                        Sql &= " tax_status_id=@tax_status_id"
                        Sql &= " WHERE ivi_id=@id"
                    Else
                        Sql = "DELETE FROM INVITEMS WHERE ivi_id=" & r.Item("id")
                    End If
                Else
                    Sql = ""
                End If
                If Sql.Length > 0 Then
                    Sql = Sql.Replace("@quantity", Me.Database.Escape(r.Item("quantity")))
                    Sql = Sql.Replace("@part_no", Me.Database.Escape(PartNo))
                    Sql = Sql.Replace("@serial_no", Me.Database.Escape(r.Item("serial_no")))
                    Sql = Sql.Replace("@description", Me.Database.Escape(r.Item("description")))
                    If Me.IsCreditMemo Then
                        Sql = Sql.Replace("@unit_price", IIf(r.Item("unit_price") Is DBNull.Value, 0, r.Item("unit_price")) * -1)
                        Sql = Sql.Replace("@unit_cost", IIf(r.Item("unit_cost") Is DBNull.Value, 0, r.Item("unit_cost")) * -1)
                    Else
                        Sql = Sql.Replace("@unit_price", IIf(r.Item("unit_price") Is DBNull.Value, 0, r.Item("unit_price")))
                        Sql = Sql.Replace("@unit_cost", IIf(r.Item("unit_cost") Is DBNull.Value, 0, r.Item("unit_cost")))
                    End If
                    Sql = Sql.Replace("@id", Me.Database.Escape(IIf(r.Item("id") Is DBNull.Value, "", r.Item("id"))))
                    Sql = Sql.Replace("@invoice_no", Me.Database.Escape(Me.InvoiceNo))
                    Sql = Sql.Replace("@tax_status_id", IIf(r.Item("tax_status_id") Is DBNull.Value, 1, r.Item("tax_status_id")))
                    Me.Database.Execute(Sql)
                    If Not Me.Database.LastQuery.Successful Then
                        Dim Err As String = Me.Database.LastQuery.ErrorMsg
                    End If
                End If
            Next
            Me.SaveGeneralLedger()
            Me.SaveCommissions()
            RaiseEvent Saved(Me)
            ' Open
            Me.Open(Me.InvoiceNo)
        Else
            Throw New Exception(Me.Database.LastQuery.ErrorMsg)
        End If

    End Sub

    Private Sub SaveGeneralLedger()
        Dim Row As DataRow
        Dim Params As New Collection
        For Each Row In Me.GeneralLedger.Rows
            Dim Sql As String = ""
            ' If Null amount, set it to zero
            If Row.Item("amount") Is DBNull.Value Then
                Row.Item("amount") = 0
            End If
            ' Decide what to do
            If Row.RowState = DataRowState.Added Then
                ' Add row if amount is over zero
                If Row.Item("amount") > 0 Or Row.Item("amount") < 0 Then
                    Sql &= "INSERT INTO invoice_gl (invoice_no, account_no, employee, amount,"
                    Sql &= " date_created, date_last_updated, created_by, last_updated_by)"
                    Sql &= " VALUES (@ino, @ano, @emp, " & Me.Database.ToCurrency("@amt") & ", @date, @date, @user, @user)"
                    Sql = Sql.Replace("@ino", Me.Database.Escape(Me.InvoiceNo))
                    Sql = Sql.Replace("@ano", Me.Database.Escape(Row.Item("account_no")))
                    If Me.IsCreditMemo Then
                        Sql = Sql.Replace("@amt", Me.Database.Escape(Row.Item("amount") * -1))
                    Else
                        Sql = Sql.Replace("@amt", Me.Database.Escape(Row.Item("amount")))
                    End If
                    Sql = Sql.Replace("@emp", Me.Database.Escape(Me.IsNull(Row.Item("employee"), "")))
                    Sql = Sql.Replace("@date", Me.Database.Escape(Now))
                    Sql = Sql.Replace("@user", Me.Database.Escape(Me.LastUpdatedBy))
                    Me.Database.Execute(Sql)
                    If Not Me.Database.LastQuery.Successful Then
                        Dim Err As String = Me.Database.LastQuery.ErrorMsg
                    End If
                End If
            ElseIf Row.RowState = DataRowState.Modified Then
                If Row.Item("amount") > 0 Or Row.Item("amount") < 0 Then
                    ' Edit row if amount over zero
                    Sql &= "UPDATE invoice_gl SET "
                    Sql &= " account_no=@ano,"
                    Sql &= " employee=@emp,"
                    Sql &= " amount=@amt,"
                    Sql &= " date_last_updated=@date,"
                    Sql &= " last_updated_by=@user"
                    Sql &= " WHERE id=@id"
                    Sql = Sql.Replace("@id", Me.Database.Escape(Row.Item("id")))
                    Sql = Sql.Replace("@ano", Me.Database.Escape(Row.Item("account_no")))
                    If Me.IsCreditMemo Then
                        Sql = Sql.Replace("@amt", Me.Database.Escape(Row.Item("amount") * -1))
                    Else
                        Sql = Sql.Replace("@amt", Me.Database.Escape(Row.Item("amount")))
                    End If
                    Sql = Sql.Replace("@emp", Me.Database.Escape(Me.IsNull(Row.Item("employee"))))
                    Sql = Sql.Replace("@date", Me.Database.Escape(Now))
                    Sql = Sql.Replace("@user", Me.Database.Escape(Me.LastUpdatedBy))
                    Me.Database.Execute(Sql)
                    If Not Me.Database.LastQuery.Successful Then
                        Dim Err As String = Me.Database.LastQuery.ErrorMsg
                    End If
                ElseIf Row.Item("id") IsNot DBNull.Value Then
                    ' Delete if amount is zero
                    Sql = "DELETE FROM invoice_gl WHERE id=" & Me.Database.Escape(Row.Item("id"))
                    Me.Database.Execute(Sql)
                End If
            End If
        Next
    End Sub

    Private Sub SaveCommissions()
        Dim Row As DataRow
        Dim Params As New Collection
        For Each Row In Me.Commissions.Rows
            ' Catch null amounts
            If Row.Item("hold") Is DBNull.Value Then
                Row.Item("hold") = 0
            End If
            If Row.Item("com_amt") Is DBNull.Value Then
                Row.Item("com_amt") = 0
            End If
            If Row.Item("hold") = 0 And Row.Item("com_amt") = 0 Then
                If Row.Item("com_id") IsNot DBNull.Value Then
                    Me.Database.Execute("DELETE FROM COMMIS WHERE com_id=" & Row.Item("com_id"))
                End If
            ElseIf Row.RowState = DataRowState.Added Then
                Dim Sql As String = "INSERT INTO COMMIS (com_no, com_inv, com_amt, hold, com_user, com_chngt)"
                Sql &= " VALUES (" & Me.Database.Escape(Row.Item("com_no"))
                Sql &= ", " & Me.Database.Escape(Me.InvoiceNo) & ", "
                If Me.IsCreditMemo Then
                    Sql &= Me.Database.Escape(Row.Item("com_amt") * -1) & ", "
                    Sql &= Me.Database.Escape(Row.Item("hold") * -1) & ", "
                Else
                    Sql &= Me.Database.Escape(Row.Item("com_amt")) & ", "
                    Sql &= Me.Database.Escape(Row.Item("hold")) & ", "
                End If
                Sql &= Me.Database.Escape(Me.LastUpdatedBy) & ", " & Me.Database.Timestamp & ")"
                Me.Database.Execute(Sql)
            ElseIf Row.RowState = DataRowState.Modified Then
                Dim Sql As String = "UPDATE COMMIS"
                Sql &= " SET"
                Sql &= " com_no=" & Me.Database.Escape(Row.Item("com_no")) & ","
                If Me.IsCreditMemo Then
                    Sql &= " com_amt=" & Me.Database.Escape(Row.Item("com_amt") * -1) & ", "
                    Sql &= " hold=" & Me.Database.Escape(Row.Item("hold") * -1) & ", "
                Else
                    Sql &= " com_amt=" & Me.Database.Escape(Row.Item("com_amt")) & ", "
                    Sql &= " hold=" & Me.Database.Escape(Row.Item("hold")) & ", "
                End If
                Sql &= " com_user=" & Me.Database.Escape(Me.LastUpdatedBy) & ", "
                Sql &= " com_chngt = " & Me.Database.Timestamp & ""
                Sql &= " WHERE com_id=" & Row.Item("com_id")
                Me.Database.Execute(Sql)
            End If
        Next
    End Sub

    Public Function TermTypes() As DataTable
        Return Me.Database.GetAll("SELECT * FROM pay_status ORDER BY sort, name")
    End Function

    Public Function TaxStatuses() As DataTable
        Return Me.Database.GetAll("SELECT id, code, description, taxable FROM tax_status ORDER BY code")
    End Function


    Public Function ToGravityDocument(ByVal Template As String) As GravityDocument.gDocument
        ' Gravity SEttings
        Dim Settings As New MyCore.cSettings(Me.Database)
        ' If no template specified
        If Template.Length = 0 Then
            Dim id As Integer
            If IsCreditMemo Then
                id = Settings.GetValue("Template Credit Memo")
            Else
                id = Settings.GetValue("Template Invoice")
            End If
            Template = Me.Database.GetOne("SELECT html FROM template WHERE id=" & id)
        End If
        ' Create Gravity Document
        Dim Doc As New GravityDocument.gDocument(Me.Database.GetOne("SELECT value FROM settings WHERE property='Page Height in Pixels'"))
        Doc.LoadXml(Template)
        ' DOCUMENT SETTINGS
        Doc.FormType = GravityDocument.gDocument.FormTypes.Invoice
        Doc.ReferenceID = Me.OrderNo
        ' Get Service order
        Dim Order As cServiceOrder = Nothing
        If Me.OrderType = TypeOfOrder.Service And Me.OrderNo > 0 Then
            Order = New cServiceOrder(Me.Database)
            Order.Open(Me.OrderNo)
        End If
        ' Invoice Items
        Dim GroupSameParts As Boolean = IIf(Settings.GetValue("Invoice Part Grouping", 1) = 1, True, False)
        Dim SummarizeLabor As Boolean = IIf(Settings.GetValue("Invoice Detail", 1) = 1, True, False)
        Dim AddServiceData As Boolean = False
        Dim NewLineItems As DataTable = Me.LineItems.Clone
        ' Customer settings override?
        Dim Customer As New cCompany(Me.Database)
        Customer.Open(Me.BillTo)
        If Customer.PrintInvoiceSetting = 4 Then ' 4 = Invoice with Service Details
            AddServiceData = True
        End If
        If Customer.InvoiceLineItemSetting = 1 Then ' 0 = Default, 1=Summarize, 2=Itemize
            SummarizeLabor = True
        End If

        ' Put in variables
        ' Get Bill To
        Dim BillToAddress As String = ""
        Dim ShipToAddress As String = ""
        Dim BillToCo As cCompany
        Dim ShipToCo As cCompany
        Try
            BillToCo = New cCompany(Me.Database)
            BillToCo.Open(Me.BillTo)
            ' If bill to has a billing address use it
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
            ' If Ship To = Bill To or Ship To = 0
            If Me.BillTo = Me.ShipTo Or Me.ShipTo = "" Then
                ShipToCo = BillToCo
                ShipToAddress = BillToCo.Address1
                If BillToCo.Address2.Length > 0 Then
                    ShipToAddress &= ControlChars.CrLf & BillToCo.Address2
                End If
            Else
                ' Look up ship to company
                ShipToCo = New cCompany(Me.Database)
                ShipToCo.Open(Me.ShipTo)
                ShipToAddress = ShipToCo.Address1
                If ShipToCo.Address2.Length > 0 Then
                    ShipToAddress &= ControlChars.CrLf & ShipToCo.Address2
                End If
            End If
        Catch ex As Exception
            Throw New Exception("Error getting company info. " & ex.ToString)
        End Try
        ' Replace variables
        Dim Page As GravityDocument.gPage
        Try
            Page = Doc.GetPage(1)
        Catch ex As Exception
            Throw New Exception("Error gettint page from template. " & ex.ToString)
        End Try
        Try
            If Not Me.IsCreditMemo Then
                Page.AddVariable("%invoice_no%", Me.InvoiceNo)
                Page.AddVariable("%invoice_date%", Format(Me.DateInvoiced, "MM/dd/yy"))
                Page.AddVariable("%due_date%", Format(Me.DateDue, "MM/dd/yy"))
                Dim Terms As String = Me.Database.GetOne("SELECT name FROM pay_status WHERE id=" & Me.TermsId)
                Page.AddVariable("%terms%", IIf(Terms = Nothing, "--", Terms))
                Page.AddVariable("%contact%", Me.ContactPerson)
                Page.AddVariable("%after_30_days%", Format(Me.TotalPlusLateFee, "$0.00"))
                Page.AddVariable("%ship_to_name%", ShipToCo.Name)
                Page.AddVariable("%ship_to_address%", ShipToAddress)
                Page.AddVariable("%ship_to_city%", ShipToCo.City)
                Page.AddVariable("%ship_to_state%", ShipToCo.State)
                Page.AddVariable("%ship_to_zip%", ShipToCo.Zip)
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
                Page.AddVariable("%tax_exempt_no%", BillToCo.TaxNo)
                If BillToCo.TaxExemptThrough <> Nothing Then
                    Page.AddVariable("%tax_exempt_thru%", BillToCo.TaxExemptThrough)
                    Page.AddVariable("%tax_exempt_yn%", IIf(BillToCo.TaxExemptThrough > Now, "Y", "N"))
                Else
                    Page.AddVariable("%tax_exempt_thru%", "--")
                    Page.AddVariable("%tax_exempt_yn%", "N")
                End If
                If Me.OrderType = cInvoice.TypeOfOrder.Rental Then
                    Page.AddVariable("%type%", "Rental")
                    If Me.OrderNo > 0 Then
                        Try
                            Dim OrderDate As Date = Me.IsNull(Me.Database.GetOne("SELECT date_ordered FROM rental_order WHERE id=" & Me.OrderNo), Nothing)
                            Page.AddVariable("%order_date%", IIf(OrderDate <> Nothing, Format(OrderDate, "MM/dd/yyyy"), "--"))
                        Catch ex As Exception
                            Page.AddVariable("%order_date%", "")
                        End Try
                        Try
                            Dim CompetedDate As Date = Me.IsNull(Me.Database.GetOne("SELECT date_delivered FROM rental_order WHERE id=" & Me.OrderNo), Nothing)
                            Page.AddVariable("%date_delivered%", IIf(CompetedDate <> Nothing, Format(CompetedDate, "MM/dd/yyyy"), "--"))
                        Catch ex As Exception
                            Page.AddVariable("%date_delivered%", "")
                        End Try
                        Page.AddVariable("%service_type%", "--")
                    Else
                        Page.AddVariable("%order_date%", "")
                        Page.AddVariable("%date_delivered%", "")
                        Page.AddVariable("%job_summary%", "")
                        Page.AddVariable("%service_type%", "--")
                    End If
                ElseIf Me.OrderType = cInvoice.TypeOfOrder.Sales Then
                    Page.AddVariable("%type%", "Sales")
                    If Me.OrderNo > 0 Then
                        Try
                            Dim OrderDate As Date = Me.IsNull(Me.Database.GetOne("SELECT date_created FROM sales_order WHERE id=" & Me.OrderNo), Nothing)
                            Page.AddVariable("%order_date%", IIf(OrderDate <> Nothing, Format(OrderDate, "MM/dd/yyyy"), "--"))
                        Catch
                            Page.AddVariable("%order_date%", "")
                        End Try
                        Try
                            Dim CompetedDate As Date = Me.IsNull(Me.Database.GetOne("SELECT date_delivered FROM sales_order WHERE id=" & Me.OrderNo), Nothing)
                            Page.AddVariable("%date_delivered%", IIf(CompetedDate <> Nothing, Format(CompetedDate, "MM/dd/yyyy"), "--"))
                        Catch
                            Page.AddVariable("%date_delivered%", "")
                        End Try
                        Page.AddVariable("%service_type%", "--")
                    Else
                        Page.AddVariable("%order_date%", "")
                        Page.AddVariable("%date_delivered%", "")
                        Page.AddVariable("%job_summary%", "")
                        Page.AddVariable("%service_type%", "--")
                    End If
                Else
                    Page.AddVariable("%type%", "Service")
                    If Order IsNot Nothing Then
                        Page.AddVariable("%order_date%", IIf(Order.DateCreated <> Nothing, Format(Order.DateCreated, "MM/dd/yyyy"), "--"))
                        Page.AddVariable("%date_delivered%", IIf(Order.DateCompleted <> Nothing, Format(Order.DateCompleted, "MM/dd/yyyy"), "--"))
                        Page.AddVariable("%job_summary%", Order.JobSummary)
                        Try
                            Page.AddVariable("%service_type%", Order.ServiceOrderTypes.Select("id=" & Order.ServiceOrderType)(0).Item("name"))
                        Catch
                            Page.AddVariable("%service_type%", "--")
                        End Try
                    Else
                        Page.AddVariable("%order_date%", "")
                        Page.AddVariable("%date_delivered%", "")
                        Page.AddVariable("%job_summary%", "")
                        Page.AddVariable("%service_type%", "--")
                    End If
                End If
                Page.AddVariable("%order_no%", Me.OrderNo)
            Else
                Page.AddVariable("%credit_no%", Me.InvoiceNo)
                Page.AddVariable("%invoice_no%", Me.OrderNo)
                Page.AddVariable("%credit_date%", Format(Me.DateInvoiced, "MM/dd/yy"))
                If BillToCo.BillingName.Length > 0 Then
                    Page.AddVariable("%customer_name%", BillToCo.BillingName)
                    Page.AddVariable("%customer_address%", BillToAddress)
                    Page.AddVariable("%customer_city%", BillToCo.BillingCity)
                    Page.AddVariable("%customer_state%", BillToCo.BillingState)
                    Page.AddVariable("%customer_zip%", BillToCo.BillingZip)
                Else
                    Page.AddVariable("%customer_name%", BillToCo.Name)
                    Page.AddVariable("%customer_address%", BillToAddress)
                    Page.AddVariable("%customer_city%", BillToCo.City)
                    Page.AddVariable("%customer_state%", BillToCo.State)
                    Page.AddVariable("%customer_zip%", BillToCo.Zip)
                End If
            End If
        Catch ex As Exception
            Throw New Exception("Error adding variables (1). " & ex.ToString)
        End Try
        ' More customer settings
        Page.AddVariable("%tax_no%", BillToCo.TaxNo)
        Page.AddVariable("%ap_email%", BillToCo.APEmailAddress)
        Page.AddVariable("%blanket_po%", BillToCo.BlanketPO)
        Page.AddVariable("%our_customer_no%", BillToCo.OurCustomerNo)
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
        ' Order no
        Try
            Page.AddVariable("%customer_no%", Me.BillTo)
            Page.AddVariable("%po%", Me.CustomerPo)
            Page.AddVariable("%salesperson%", Me.Salesman)
            Page.AddVariable("%subtotal%", Format(Me.Subtotal, "$0.00"))
            Page.AddVariable("%freight%", Format(Me.Freight, "$0.00"))
            Page.AddVariable("%tax%", Format(Me.SalesTax, "$0.00"))
            Page.AddVariable("%total%", Format(Me.Total, "$0.00"))
            Page.AddVariable("%message%", Me.Memo)
            Page.AddVariable("%memo%", Me.Memo)
        Catch ex As Exception
            Throw New Exception("Error adding variables (2). " & ex.ToString)
        End Try



        Try
            If GroupSameParts Then
                Dim Hash As New Hashtable
                Dim Index As Integer = 0
                For Each Row As DataRow In Me.LineItems.Rows
                    Dim Type As Integer = IIf(Row.Item("item_type_id") Is DBNull.Value, 1, Row.Item("item_type_id"))
                    If Type <> 2 And Type <> 3 Then
                        Dim Key As String = Row.Item("part_no") & "__" & Row.Item("unit_price").ToString.Replace(".", "_")
                        If Hash.ContainsKey(Key) Then
                            NewLineItems.Rows(Hash.Item(Key)).Item("quantity") += Row.Item("quantity")
                            NewLineItems.Rows(Hash.Item(Key)).Item("serial_no") = ""
                            NewLineItems.Rows(Hash.Item(Key)).Item("ext_price") += Row.Item("quantity") * Row.Item("unit_price")
                        Else
                            Hash.Add(Key, Index)
                            Dim NewRow As DataRow = NewLineItems.NewRow
                            For i As Integer = 0 To NewLineItems.Columns.Count - 1
                                NewRow.Item(i) = Row.Item(i)
                            Next
                            NewLineItems.Rows.Add(NewRow)
                            Index += 1
                        End If
                    End If
                Next
            Else
                For Each Row As DataRow In Me.LineItems.Rows
                    Dim Type As Integer = IIf(Row.Item("item_type_id") Is DBNull.Value, 1, Row.Item("item_type_id"))
                    If Type <> 2 And Type <> 3 Then
                        Dim NewRow As DataRow = NewLineItems.NewRow
                        For i As Integer = 0 To NewLineItems.Columns.Count - 1
                            NewRow.Item(i) = Row.Item(i)
                        Next
                        NewLineItems.Rows.Add(NewRow)
                    End If
                Next
            End If
        Catch
            ' Ignore
        End Try
        Try
            If SummarizeLabor Then
                ' Line Item: Trip and Labor
                ' Get last tax status and assume all use that
                ' Dangerous?  Maybe, but we want to group these together on the printed invoice.
                Dim TaxStatus As Integer = 0
                Dim Total As Double = 0
                For Each Row As DataRow In Me.LineItems.Rows
                    Dim Type As Integer = IIf(Row.Item("item_type_id") Is DBNull.Value, 1, Row.Item("item_type_id"))
                    If Type = 2 Or Type = 3 Then
                        TaxStatus = Row.Item("tax_status_id")
                        Total += Row.Item("quantity") * Row.Item("unit_price")
                    End If
                Next
                ' If a line item was found in the trip/labor categories, tax status will be greater than zero
                If TaxStatus > 0 Then
                    Dim nr As DataRow = NewLineItems.NewRow
                    nr.Item("quantity") = 1
                    nr.Item("part_no") = ""
                    nr.Item("serial_no") = ""
                    If Me.OrderType = TypeOfOrder.Service Then
                        nr.Item("description") = Settings.GetValue("Service Order Grouped Labor Description", "Trip/Labor, Service Order #%order_no%").Replace("%order_no%", Me.OrderNo)
                    ElseIf Me.OrderType = TypeOfOrder.Rental Then
                        nr.Item("description") = Settings.GetValue("Rental Order Grouped Labor Description", "Delivery/Setup, Rental Order #%order_no%").Replace("%order_no%", Me.OrderNo)
                    Else
                        nr.Item("description") = Settings.GetValue("Sales Order Grouped Labor Description", "Delivery/Setup, Sales Order #%order_no%").Replace("%order_no%", Me.OrderNo)
                    End If
                    nr.Item("unit_price") = Total
                    nr.Item("ext_price") = Total
                    nr.Item("unit_cost") = 0
                    nr.Item("tax_status_id") = TaxStatus
                    NewLineItems.Rows.Add(nr)
                End If
            Else
                For Each Row As DataRow In Me.LineItems.Rows
                    Dim Type As Integer = IIf(Row.Item("item_type_id") Is DBNull.Value, 1, Row.Item("item_type_id"))
                    If Type = 2 Or Type = 3 Then
                        Dim nr As DataRow = NewLineItems.NewRow
                        For i As Integer = 0 To NewLineItems.Columns.Count - 1
                            nr.Item(i) = Row.Item(i)
                        Next
                        NewLineItems.Rows.Add(nr)
                    End If
                Next
            End If
        Catch ex As Exception
            Throw New Exception("Error generating line items. " & ex.ToString)
        End Try

        ' Add service details line items?
        If Order IsNot Nothing And AddServiceData Then
            Dim Sql As String = "SELECT wo.equipment_id, d.dep_manuf, d.dep_mod, d.dep_ser, d.dep_assno FROM work_order wo INNER JOIN DEPREC d ON d.dep_id=wo.equipment_id"
            Sql &= " WHERE service_order_id=" & Order.OrderNo & " AND equipment_id > 0"
            Dim Equipment As DataTable = Me.Database.GetAll(Sql)
            For Each Row As DataRow In Equipment.Rows
                Dim nr As DataRow = NewLineItems.NewRow
                Dim Desc As String = ""
                If Row.Item("dep_manuf") IsNot DBNull.Value Then
                    Desc &= Row.Item("dep_manuf") & " "
                End If
                If Row.Item("dep_mod") IsNot DBNull.Value Then
                    Desc &= Row.Item("dep_mod")
                End If
                If Row.Item("dep_assno") IsNot DBNull.Value Then
                    Desc &= " - " & Row.Item("dep_assno")
                End If
                nr.Item("description") = Desc
                If Row.Item("dep_ser") IsNot DBNull.Value Then
                    nr.Item("serial_no") = Row.Item("dep_ser")
                End If
                NewLineItems.Rows.Add(nr)
            Next
        End If


        ' Add Line Items to Template
        Try
            If Me.IsCreditMemo Then
                Page.GetTableBySource("line_items").Table.Data = NewLineItems
            Else
                Page.GetTableBySource("invoice_items").Table.Data = NewLineItems
            End If
        Catch ex As Exception
            'Throw New Exception("Error trying to print line items. " & ex.ToString)
            ' If we get this error, it probably means that the table did not exist in the template
        End Try

        ' Return Value
        Return Doc
    End Function

    Public Sub IncrementNextNumber()
        Me.Database.Execute("UPDATE next_number SET number=number+1 WHERE name='invoice'")
    End Sub

    Public Function GetNextNumber() As Integer
        Return Me.Database.GetOne("SELECT number FROM next_number WHERE name='invoice'")
    End Function


End Class
