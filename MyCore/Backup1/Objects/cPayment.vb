Imports MyCore.Data

Public Class cPayment

    Dim Database As MyCore.Data.EasySql

    Dim _CustomerName As String = ""

    Public RecordNo As Integer = 0
    Public CustomerNo As String = ""
    Public PaymentDate As Date = Today
    Public Amount As Double = 0
    Public CheckNo As String = ""
    Public Method As Integer = 0
    Public Note As String = ""

    Public AcctRef As String = Nothing
    Public AcctUpload As Boolean = Nothing
    Public CustomerRef As String = Nothing
    Public MethodRef As String = Nothing

    Public LastUpdatedBy As String = ""
    Public CreatedBy As String = ""

    Public GeneralLedgerEntries As DataTable

    Public Event Reload()
    Public Event Saved(ByVal Payment As cPayment)

    Public ReadOnly Property CustomerName() As String
        Get
            Return Me._CustomerName
        End Get
    End Property

    Public ReadOnly Property PaymentMethods() As DataTable
        Get
            Return Me.Database.GetAll("SELECT * FROM payment_method")
        End Get
    End Property

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
        ' Make General Ledger Table
        Dim Table As New DataTable
        Table.Columns.Add("pgl_id")
        Table.Columns.Add("pgl_acc")
        Table.Columns.Add("pgl_amt")
        Table.Columns.Add("pgl_ino")
        Table.Columns.Add("pgl_user")
        Table.Columns.Add("pgl_chngt")
        Me.GeneralLedgerEntries = Table
    End Sub

    Public Sub Open(ByVal no As Integer)
        Dim Sql As String = "SELECT p.*, a.acct_ref AS customer_ref, pm.acct_ref AS pm_ref, a.cst_name"
        Sql &= " FROM [PAYMENTS] p"
        Sql &= " LEFT OUTER JOIN ADDRESS a ON p.pmt_no=a.cst_no"
        Sql &= " LEFT OUTER JOIN payment_method pm ON p.payment_method=pm.id"
        Sql &= " WHERE pmt_recno=" & no
        Dim Row As DataRow = Me.Database.GetRow(Sql)
        If Me.Database.LastQuery.RowsReturned = 1 Then
            Me.RecordNo = no
            Me.PaymentDate = Row.Item("pmt_date")
            Me.CustomerNo = Row.Item("pmt_no")
            Me.Amount = Row.Item("pmt_amt")
            Me.CheckNo = Row.Item("pmt_chck")
            Me.Method = Row.Item("payment_method")
            Me.Note = Row.Item("note")
            Me.LastUpdatedBy = Row.Item("pmt_user")
            Me.CreatedBy = Row.Item("pmt_user")
            Me.AcctRef = IIf(Row.Item("acct_ref") Is DBNull.Value, Nothing, Row.Item("acct_ref"))
            Me.AcctUpload = IIf(Row.Item("acct_upload") Is DBNull.Value, Nothing, Row.Item("acct_upload"))
            Me.CustomerRef = IIf(Row.Item("customer_ref") Is DBNull.Value, Nothing, Row.Item("customer_ref"))
            Me.MethodRef = IIf(Row.Item("pm_ref") Is DBNull.Value, Nothing, Row.Item("pm_ref"))
            Me._CustomerName = Row.Item("cst_name")
            ' Get line items
            Sql = "SELECT pgl_id, pgl_acc, pgl_amt, pgl_ino, pgl_user, pgl_chngt"
            Sql &= " FROM PYMNTGLD"
            Sql &= " WHERE pgl_recno=" & no
            Me.GeneralLedgerEntries = Me.Database.GetAll(Sql)
            ' Raise Reload
            RaiseEvent Reload()
        Else
            Throw New Exception("No payment with record number " & no & " was found.")
        End If
    End Sub

    Public Function IsDuplicateCheckNo(ByVal CustomerNo As String, ByVal CheckNo As String) As Boolean
        Dim Sql As String = "SELECT * FROM PAYMENTS WHERE pmt_chck=" & Me.Database.Escape(CheckNo)
        Sql &= " AND pmt_no=" & Me.Database.Escape(CustomerNo)
        If Me.RecordNo > 0 Then
            Sql &= " AND pmt_recno <> " & Me.RecordNo
        End If
        Dim Table As DataTable = Me.Database.GetAll(Sql)
        If Table.Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function Save() As Boolean
        If Me.DuplicateCheckNo Then
            Throw New Exception("Duplicate check number.")
            Return False
        End If
        If Me.Amount < 0 Then
            Throw New Exception("Cannot have a payment with a negative amount.")
            Return False
        End If
        If Me.RecordNo = 0 Then
            Me.SaveNew()
            If Me.Database.LastQuery.Successful Then
                Me.SaveLineItems()
                Me.Open(Me.RecordNo)
                RaiseEvent Saved(Me)
                Return True
            End If
        Else
            Me.SaveExisting()
            If Me.Database.LastQuery.Successful Then
                Me.SaveLineItems()
                RaiseEvent Saved(Me)
                Me.Open(Me.RecordNo)
                Return True
            End If
        End If
        Throw New Exception(Me.Database.LastQuery.ErrorMsg)
        Return False
    End Function

    Private Sub SaveExisting()
        Dim Sql As String = "UPDATE PAYMENTS SET"
        Sql &= " pmt_no=" & Me.Database.Escape(Me.CustomerNo) & ", "
        Sql &= " pmt_date=" & Me.Database.Escape(Me.PaymentDate) & ", "
        Sql &= " pmt_chck=" & Me.Database.Escape(Me.CheckNo) & ", "
        Sql &= " pmt_amt=" & Me.Database.Escape(Me.Amount) & ", "
        Sql &= " pmt_chngt=" & Me.Database.Timestamp & ", "
        Sql &= " pmt_user=" & Me.Database.Escape(Me.LastUpdatedBy) & ", "
        Sql &= " payment_method=" & Me.Database.Escape(Me.Method) & ", "
        Sql &= " note=" & Me.Database.Escape(Me.Note)
        Sql &= " WHERE pmt_recno=" & Me.RecordNo
        Me.Database.Execute(Sql)
    End Sub

    Private Sub SaveNew()
        ' Get next record number
        Dim RecordNo As Integer = 1
        Try
            RecordNo = Me.GetNextNumber
        Catch
            RecordNo = 1
        End Try
        ' Query
        Dim Sql As String = "INSERT INTO PAYMENTS (pmt_no, pmt_date, pmt_chck, pmt_amt, pmt_chngt, pmt_user, "
        Sql &= " payment_method, note, pmt_recno)"
        Sql &= " VALUES (@customer, @payment_date, @check_no, @amount, " & Me.Database.Timestamp & ", @windows_user, "
        Sql &= " @method, @note, @record_no)"
        Sql = Sql.Replace("@record_no", Me.Database.Escape(RecordNo))
        Sql = Sql.Replace("@customer", Me.Database.Escape(Me.CustomerNo))
        Sql = Sql.Replace("@payment_date", Me.Database.Escape(Me.PaymentDate))
        Sql = Sql.Replace("@check_no", Me.Database.Escape(Me.CheckNo))
        Sql = Sql.Replace("@amount", Me.Database.Escape(Me.Amount))
        Sql = Sql.Replace("@windows_user", Me.Database.Escape(Me.LastUpdatedBy))
        Sql = Sql.Replace("@method", Me.Database.Escape(Me.Method))
        Sql = Sql.Replace("@note", Me.Database.Escape(Me.Note))
        Me.Database.Execute(Sql)
        If Me.Database.LastQuery.Successful Then
            Me.RecordNo = RecordNo
            Me.IncrementNextNumber()
        End If
    End Sub

    Private Sub SaveLineItems()
        Dim Invoices As New Hashtable
        Dim InvNo As String
        Dim Row As DataRow
        Dim InvCount As Integer = 0
        Dim Params As New Collection
        For Each Row In Me.GeneralLedgerEntries.Rows
            Params = New Collection
            If Row.RowState = DataRowState.Added Then
                Dim Sql As String = "INSERT INTO PYMNTGLD (pgl_recno, pgl_acc, pgl_amt, pgl_ino, pgl_user, pgl_chngt)"
                Sql &= " VALUES (@record_no, @account, " & Me.Database.ToCurrency("@amount") & ", @invoice, @windows_user, " & Me.Database.Timestamp & ")"
                Sql = Sql.Replace("@record_no", Me.RecordNo)
                Sql = Sql.Replace("@account", Me.Database.Escape(Row.Item("pgl_acc")))
                Sql = Sql.Replace("@invoice", Me.Database.Escape(Row.Item("pgl_ino")))
                Sql = Sql.Replace("@amount", Me.Database.Escape(Row.Item("pgl_amt")))
                Sql = Sql.Replace("@windows_user", Me.Database.Escape(Me.LastUpdatedBy))
                Me.Database.Execute(Sql)
                If Not Me.Database.LastQuery.Successful Then
                    Dim Err As String = Me.Database.LastQuery.ErrorMsg
                End If
            ElseIf Row.RowState = DataRowState.Modified Then
                If Row.Item("pgl_id") IsNot DBNull.Value And Row.Item("pgl_amt") IsNot DBNull.Value Then
                    If Row.Item("pgl_amt") = 0 Then
                        ' Delete it
                        Me.Database.Execute("DELETE FROM PYMNTGLD WHERE pgl_id=" & Row.Item("pgl_id"))
                    Else
                        ' Edit it
                        Dim Sql As String = "UPDATE PYMNTGLD "
                        Sql &= " SET pgl_acc=@account, pgl_ino=@invoice, pgl_amt=@amount, pgl_user=@windows_user,"
                        Sql &= " pgl_chngt = " & Me.Database.Timestamp & ""
                        Sql &= " WHERE pgl_id=@item_id"
                        Sql = Sql.Replace("@record_no", Me.RecordNo)
                        Sql = Sql.Replace("@account", Me.Database.Escape(Row.Item("pgl_acc")))
                        Sql = Sql.Replace("@invoice", Me.Database.Escape(Row.Item("pgl_ino")))
                        Sql = Sql.Replace("@amount", Me.Database.Escape(Row.Item("pgl_amt")))
                        Sql = Sql.Replace("@windows_user", Me.Database.Escape(Me.LastUpdatedBy))
                        Sql = Sql.Replace("@item_id", Me.Database.Escape(Row.Item("pgl_id")))
                        Me.Database.Execute(Sql)
                    End If
                End If
            End If
            ' Add invoice to hashtable
            If Not Invoices.ContainsValue(Row.Item("pgl_ino")) And Row.Item("pgl_ino").ToString.Length > 0 Then
                Invoices.Add(InvCount, Row.Item("pgl_ino"))
                InvCount += 1
            End If
        Next
        ' Look and see if invoices in this payment are now paid off... if so mark them as paid 
        For i As Integer = 0 To InvCount - 1
            Dim Invoice As New cInvoice(Me.Database)
            Invoice.Open(Invoices(i))
            ' Need to keep it unpaid if they overpay so that it shows up on statements, so only if it = 0
            If Invoice.AmountPaid = Invoice.Total Then
                Invoice.Paid = True
                Invoice.Save()
            ElseIf Invoice.Paid Then
                Invoice.Paid = False
                Invoice.Save()
            End If
        Next
    End Sub

    Private Function DuplicateCheckNo() As Boolean
        Dim Sql As String = "SELETE COUNT(pmt_id) WHERE pmt_no='" & Me.CustomerNo & "' AND pmt_chck='" & Me.CheckNo & "'"
        Dim Count As Integer = 0
        Try
            Count = Me.Database.GetOne(Sql)
        Catch
            Return False
        End Try
        If Me.RecordNo > 0 And Count > 1 Then
            Return True
        ElseIf Me.RecordNo = 0 And Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub IncrementNextNumber()
        Me.Database.Execute("UPDATE next_number SET number=number+1 WHERE name='payment'")
    End Sub

    Public Function GetNextNumber() As Integer
        Return Me.Database.GetOne("SELECT number FROM next_number WHERE name='payment'")
    End Function

End Class
