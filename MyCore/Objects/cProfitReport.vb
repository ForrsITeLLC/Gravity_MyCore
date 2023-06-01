Public Class cProfitReport

    Dim Database As MyCore.Data.EasySql
    Dim Settings As cSettings

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
        Me.Settings = New cSettings(Me.Database)
    End Sub

    Public Function StatementOfCondition(ByVal TemplateId As Integer, ByVal StartDate As Date, ByVal EndDate As Date) As GravityDocument.gDocument
        ' Template
        Dim Template As String = Me.Database.GetOne("SELECT html FROM template WHERE id=" & TemplateId)
        ' Create Document
        Dim Doc As New GravityDocument.gDocument(Settings.GetValue("Page Height in Pixels", 912))
        ' Get Offices
        Dim Offices As DataTable = Me.Database.GetAll("SELECT number, name FROM office ORDER BY sort, name")
        ' Work order pages
        For Each Row As DataRow In Offices.Rows
            Dim PageDoc As New GravityDocument.gDocument(Settings.GetValue("Page Height in Pixels", 912))
            Dim Page As GravityDocument.gPage = PageDoc.AddPageFromXml(Template)
            ' Get Some Values
            Dim AR As DataTable = Me.GetAccountsReceivable(Row.Item("number"), StartDate, EndDate)
            Dim Revenue As Double = Me.TotalTable(AR)
            Dim Costs As Double = Me.GetCosts(Row.Item("number"), StartDate, EndDate)
            Dim Commissions As Double = Me.GetCommissionsTotal(Row.Item("number"), StartDate, EndDate)
            Dim Payments As Double = Me.GetPayments(Row.Item("number"), StartDate, EndDate)
            Dim POs As Double = Me.GetOutstandingPOsTotal(Row.Item("number"))
            Dim TotalBalance As Double = Me.GetOutstandingInvoices(Row.Item("number"))
            ' Put in variables
            Page.AddVariable("%office_number%", Row.Item("number"))
            Page.AddVariable("%office_name%", Row.Item("name"))
            Page.AddVariable("%start%", StartDate)
            Page.AddVariable("%end%", EndDate)
            ' Populate accounts receivable
            For i As Integer = 0 To AR.Rows.Count - 1
                Page.AddVariable("%ar_" & AR.Rows(i).Item("category").ToString.Replace(" ", "_") & "%", Format(AR.Rows(i).Item("total"), "c"))
            Next
            Page.AddVariable("%Total_Revenue%", Format(Revenue, "c"))
            ' Payments
            Page.AddVariable("%payments%", Format(Payments, "c"))
            ' Costs
            Page.AddVariable("%total_costs%", Format(Costs, "c"))
            ' Outstanding POs
            Page.AddVariable("%outstanding_pos%", Format(POs, "c"))
            ' Outstanding Invoices
            Page.AddVariable("%outstanding_invoices%", Format(TotalBalance, "c"))
            ' Commissions Payable
            Page.AddVariable("%commissions%", Format(Commissions, "c"))
            ' Gross profit
            Page.AddVariable("%Total_Gross_Profit%", Format(Revenue - Costs - Commissions, "c"))
            ' Add to main document
            Doc.AddPage(PageDoc.Pages(1))
        Next
        Return Doc
    End Function

    Private Function GetOutstandingInvoices(ByVal Office As String) As Double
        Dim Sql As String = ""
        Sql = "SELECT Balance = SUM(ISNULL(iamt, 0) - ISNULL(pamt, 0))"
        Sql &= " FROM [INVOICE] inv"
        Sql &= " LEFT OUTER JOIN (SELECT invoice_no, iamt = SUM(amount) FROM invoice_gl"
        Sql &= " GROUP BY invoice_no) igl ON inv.inv_ino = igl.invoice_no"
        Sql &= " LEFT OUTER JOIN (SELECT pgl_ino, pamt = SUM(ISNULL(pgl_amt,0)) FROM PYMNTGLD"
        Sql &= " GROUP BY pgl_ino) pgl ON inv.inv_ino = pgl.pgl_ino"
        Sql &= " WHERE inv.voided=0 AND inv.refund=0 AND inv.inv_paid = 0 AND inv.inv_off=@Office"
        Sql &= " GROUP BY inv_off"
        Sql = Sql.Replace("@Office", Me.Database.Escape(Office))
        Return Me.Database.GetOne(Sql)
    End Function

    Private Function GetPayments(ByVal Office As String, ByVal StartDate As Date, ByVal EndDate As Date) As Double
        Dim Sql As String = ""
        Sql &= " SELECT SUM(ISNULL(pamt, 0))"
        Sql &= " FROM PAYMENTS pmt"
        Sql &= " LEFT OUTER JOIN (SELECT pgl_recno, pamt = SUM(ISNULL(pgl_amt, 0))"
        Sql &= " FROM PYMNTGLD pgl, INVOICE inv"
        Sql &= " WHERE pgl.pgl_ino=inv.inv_ino AND inv_off=@Office"
        Sql &= " GROUP BY pgl_recno) pgl ON pmt.pmt_recno=pgl.pgl_recno"
        Sql &= " WHERE pmt_date BETWEEN @start AND @end"
        Sql = Sql.Replace("@Office", Me.Database.Escape(Office))
        Sql = Sql.Replace("@start", Me.Database.Escape(StartDate))
        Sql = Sql.Replace("@end", Me.Database.Escape(EndDate))
        Return Me.Database.GetOne(Sql)
    End Function

    Private Function GetAccountsReceivable(ByVal Office As Integer, ByVal StartDate As Date, ByVal EndDate As Date) As DataTable
        Dim Sql As String = "SELECT office, cat.name AS category, ISNULL(total, 0) AS total"
        Sql &= " FROM gl_account_category cat"
        Sql &= " LEFT OUTER JOIN ("
        Sql &= " SELECT inv_off AS office, account_category_id, "
        Sql &= " SUM(ISNULL(igl.amount, 0)) AS total"
        Sql &= " FROM INVOICE"
        Sql &= " INNER JOIN invoice_gl igl ON igl.invoice_no=inv_ino"
        Sql &= " INNER JOIN gl_account acc ON igl.account_no=acc.account_no"
        Sql &= " WHERE (inv_date BETWEEN @date_start AND @date_end)"
        Sql &= " AND inv_off=@office AND voided=0"
        Sql &= " GROUP BY inv_off, account_category_id"
        Sql &= " ) gl ON cat.id=gl.account_category_id"
        Sql &= " ORDER BY office, category"
        Sql = Sql.Replace("@date_start", Me.Database.Escape(StartDate))
        Sql = Sql.Replace("@date_end", Me.Database.Escape(EndDate))
        Sql = Sql.Replace("@office", Me.Database.Escape(Office))
        Dim Table As DataTable = Me.Database.GetAll(Sql)
        If Not Me.Database.LastQuery.Successful Then
            Throw New Exception(Me.Database.LastQuery.ErrorMsg)
        End If
        Return Table
    End Function

    Private Function GetGrossProfit(ByVal Office As Integer, ByVal StartDate As Date, ByVal EndDate As Date) As Double
        Dim Sql As String = ""
        Sql &= "SELECT total=SUM(ISNULL(invoices.gp, 0)) FROM ("
        Sql &= " SELECT gp = CASE refund "
        Sql &= " WHEN 0 THEN (ISNULL(items.total, 0) + ISNULL(inv_shch, 0)) * (1 + (ISNULL(inv_taxp, 0) / 100)) - ISNULL(com.total, 0) - ISNULL(items.cost, 0)"
        Sql &= " ELSE ((ISNULL(items.total, 0) + ISNULL(inv_shch, 0)) * (1 + (ISNULL(inv_taxp, 0) / 100)) - ISNULL(com.total, 0) - ISNULL(items.cost, 0)) *-1 END"
        Sql &= " FROM INVOICE"
        Sql &= " INNER JOIN (SELECT ino=MAX(ivi_ino), total=SUM(ISNULL(ivi_qua, 0) * ISNULL(ivi_cost, 0)), cost=SUM(ISNULL(ivi_qua, 0) * ISNULL(ivi_ocost, 0))"
        Sql &= " FROM INVITEMS GROUP BY ivi_ino) items ON inv_ino=items.ino"
        Sql &= " LEFT OUTER JOIN (SELECT ino=MAX(com_inv), total=SUM(ISNULL(com_amt, 0)) FROM COMMIS GROUP BY com_inv) com ON com.ino=inv_ino"
        Sql &= " WHERE voided=0 AND inv_date BETWEEN @start AND @end"
        Sql &= " AND inv_off=@Office) invoices"
        Sql = Sql.Replace("@Office", Me.Database.Escape(Office))
        Sql = Sql.Replace("@start", Me.Database.Escape(StartDate))
        Sql = Sql.Replace("@end", Me.Database.Escape(EndDate))
        Return Me.Database.GetOne(Sql)
    End Function

    Private Function GetCosts(ByVal Office As Integer, ByVal StartDate As Date, ByVal EndDate As Date) As Double
        Dim Sql As String = ""
        Sql &= "SELECT total=SUM(costs) FROM ("
        Sql &= " SELECT costs = CASE refund WHEN 0 THEN ISNULL(items.cost, 0) ELSE ISNULL(items.cost, 0) * -1 END"
        Sql &= " FROM INVOICE"
        Sql &= " INNER JOIN (SELECT ivi_ino, cost=SUM(ISNULL(ivi_qua, 0) * ISNULL(ivi_ocost, 0))"
        Sql &= " FROM INVITEMS GROUP BY ivi_ino) items ON inv_ino=items.ivi_ino"
        Sql &= " WHERE voided=0 AND inv_date BETWEEN @start AND @end"
        Sql &= " AND inv_off=@Office) invoices"
        Sql = Sql.Replace("@Office", Me.Database.Escape(Office))
        Sql = Sql.Replace("@start", Me.Database.Escape(StartDate))
        Sql = Sql.Replace("@end", Me.Database.Escape(EndDate))
        Return Me.Database.GetOne(Sql)
    End Function

    Private Function GetOutstandingPOsTotal(ByVal Office As Integer) As Double
        Dim Sql As String = "SELECT ISNULL(SUM((quantity-received)*unit_price*(1+discount/100)), 0) AS total"
        Sql &= " FROM purchase_order_item poi"
        Sql &= " INNER JOIN purchase_order po ON poi.po_no=po.po_no"
        Sql &= " WHERE quantity > received AND po.office=@office AND rma=0"
        Sql = Sql.Replace("@office", Me.Database.Escape(Office))
        Return Me.Database.GetOne(Sql)
    End Function

    Private Function GetCommissionsTotal(ByVal Office As Integer, ByVal StartDate As Date, ByVal EndDate As Date) As Double
        Dim Sql As String = ""
        Sql &= " SELECT total= CASE refund WHEN 0 THEN ISNULL(SUM(com_amt), 0) ELSE ISNULL(SUM(com_amt), 0) * -1 END"
        Sql &= " FROM COMMIS"
        Sql &= " INNER JOIN INVOICE ON inv_ino=com_inv"
        Sql &= " WHERE (inv_date BETWEEN @date_start AND @date_end) AND inv_off=@office AND voided=0"
        Sql &= " GROUP BY inv_off, refund"
        Sql = Sql.Replace("@date_start", Me.Database.Escape(StartDate))
        Sql = Sql.Replace("@date_end", Me.Database.Escape(EndDate))
        Sql = Sql.Replace("@office", Me.Database.Escape(Office))
        Return Me.Database.GetOne(Sql)
    End Function

    Private Function TotalTable(ByVal Table As DataTable) As Double
        Dim r As DataRow
        Dim total As Double = 0
        For Each r In Table.Rows
            total += r.Item("total")
        Next
        Return total
    End Function


End Class
