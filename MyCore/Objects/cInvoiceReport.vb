Public Class cInvoiceReport

    Dim Database As MyCore.Data.EasySql

    Public DateStart As DateTime = Nothing
    Public DateEnd As DateTime = Nothing
    Public Show As ShowValues = ShowValues.Both
    Public PrintStatus As PrintStatusValues = PrintStatusValues.Either
    Public MinimumDaysOld As Integer = 0
    Public PaidStatus As PaidStatusValues = PaidStatusValues.Either
    Public Office As String = ""

    Public Enum ShowValues
        Invoices = 0
        Credits = 1
        Both = 2
    End Enum

    Public Enum PrintStatusValues
        Printed = 0
        Unprinted = 1
        Either = 2
    End Enum

    Public Enum PaidStatusValues
        Paid = 0
        Unpaid = 1
        Either = 2
    End Enum

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
    End Sub

    Public Function GetSql() As String
        Dim Sql As String = ""
        Sql &= " SELECT * FROM"
        Sql &= " ("
        Sql &= " SELECT "
        Sql &= " inv_ino AS [invoice_no], [inv_date] AS [date], ROUND([total], 2) AS total, ROUND([balance], 2) AS balance,"
        Sql &= " [cst_name] AS [customer_name], ([cst_city] + ', ' + [cst_state]) AS [customer_city],"
        Sql &= " [cst_no] AS [customer_no], inv_off AS office, type = CASE refund WHEN 0 THEN 'Invoice' ELSE 'Credit' END,"
        Sql &= " i.*"
        Sql &= " FROM INVOICE AS i"
        Sql &= " LEFT JOIN ("
        Sql &= " SELECT  [ivi_ino], "
        Sql &= " ISNULL(SUM([ivi_qua] * [ivi_cost]), 0)  + ISNULL(AVG([inv_shch]), 0)  "
        Sql &= " + MAX(sales_tax) AS [total],"
        Sql &= " (ISNULL(SUM([ivi_qua] * [ivi_cost]), 0)  + ISNULL(AVG([inv_shch]), 0)  "
        Sql &= " + MAX(sales_tax)) -"
        Sql &= " ISNULL((SELECT SUM(pgl_amt) FROM PYMNTGLD WHERE pgl_ino=i.ivi_ino GROUP BY pgl_ino), 0)"
        Sql &= " AS [balance]"
        Sql &= " FROM INVITEMS i, INVOICE "
        Sql &= " WHERE [ivi_ino] IN (SELECT [inv_ino] FROM INVOICE) "
        Sql &= " AND [ivi_ino]=[inv_ino] GROUP BY [ivi_ino]"
        Sql &= " ) AS d ON i.inv_ino=d.ivi_ino"
        Sql &= " LEFT JOIN ADDRESS ON [inv_no]=[cst_no]"
        Sql &= " WHERE voided=0"
        If Me.MinimumDaysOld > 0 Then
            Sql &= " AND " & Me.Database.DiffDays("inv_date", Me.Database.Timestamp) & " >= " & Me.MinimumDaysOld
        End If
        If Me.PrintStatus = PrintStatusValues.Unprinted Then
            Sql &= " AND inv_prnt='Y'"
        ElseIf Me.PrintStatus = PrintStatusValues.Printed Then
            Sql &= " AND inv_prnt='N'"
        End If
        If Me.Show = ShowValues.Invoices Then
            Sql &= " AND refund=0"
        ElseIf Me.Show = ShowValues.Credits Then
            Sql &= " AND refund=1"
        End If
        If Me.PaidStatus = PaidStatusValues.Unpaid Then
            Sql &= " AND inv_paid=0"
        ElseIf Me.PaidStatus = PaidStatusValues.Paid Then
            Sql &= " AND inv_paid=1"
        End If
        If Me.DateStart <> Nothing And Me.DateEnd <> Nothing Then
            Sql &= " AND (inv_date BETWEEN " & Me.Database.Escape(Me.DateStart) & " AND " & Me.Database.Escape(Me.DateEnd) & ")"
        ElseIf Me.DateStart <> Nothing Then
            Sql &= " AND inv_date >= " & Me.Database.Escape(Me.DateStart)
        ElseIf Me.DateEnd <> Nothing Then
            Sql &= " AND inv_date >= " & Me.Database.Escape(Me.DateStart)
        End If
        If Me.Office.Length > 0 Then
            Sql &= " AND inv_off = " & Me.Database.Escape(Me.Office)
        End If
        Sql &= " ) AS [results]"
        Return Sql
    End Function

    Public Function GetDataTable() As DataTable
        Return Me.Database.GetAll(Me.GetSql)
    End Function

End Class
