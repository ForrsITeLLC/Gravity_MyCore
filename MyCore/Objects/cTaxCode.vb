Public Class cTaxCode

    Dim Database As MyCore.Data.EasySql

    Dim _Id As Integer = 0
    Dim _Authorities As DataTable

    Public Name As String = ""
    Public FreightMethod As FreightTaxMethod

    Public Enum FreightTaxMethod
        Never = 0
        WhenSomethingElseTaxed = 1
        Always = 2
    End Enum


    Public ReadOnly Property BasePercent() As Double
        Get
            Dim Sql As String = "SELECT SUM(rate) FROM tax_authority" & _
            " JOIN tax_code_item ON tax_code_item.tax_authority_id=tax_authority.id" & _
            " WHERE tax_code_item.tax_code_id=" & Me._Id
            Try
                Dim p As Double = Me.Database.GetOne(Sql)
                Return p
            Catch
                Return 0
            End Try
        End Get
    End Property

    Public Structure LineItem
        Dim Quantity As Double
        Dim Amount As Double
        Dim Taxable As Boolean
    End Structure

    Public ReadOnly Property TaxAuthorities() As DataTable
        Get
            Return Me._Authorities
        End Get
    End Property

    Public ReadOnly Property Id() As Integer
        Get
            Return Me._Id
        End Get
    End Property

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
    End Sub

    Public Sub Open(ByVal TaxCodeId As Integer)
        Dim Code As DataRow = Me.Database.GetRow("SELECT name, freight_method FROM tax_code WHERE id=" & TaxCodeId)
        If Me.Database.LastQuery.RowsReturned = 1 Then
            ' Get Code
            Me._Id = TaxCodeId
            Me.Name = Code.Item("name")
            Me.FreightMethod = Code.Item("freight_method")
            ' Get Tax Authorities
            Dim Sql As String = ""
            Sql &= "SELECT a.id, a.name, a.rate, a.formula_base, a.dollar_limit, a.above_limit_rate, a.gl_account"
            Sql &= " FROM tax_authority a WHERE a.id IN (SELECT tax_authority_id FROM tax_code_item WHERE tax_code_id=" & TaxCodeId & ")"
            Me._Authorities = Me.Database.GetAll(Sql)
            If Not Me.Database.LastQuery.Successful Then
                Dim er As String = Me.Database.LastQuery.ErrorMsg
            End If
        Else
            Dim er As String = Me.Database.LastQuery.ErrorMsg
            Throw New Exception("Tax code not found.")
        End If
    End Sub

    Public Function SetByCompany(ByVal CustomerNo As String) As Integer
        If CustomerNo.Length > 0 Then
            Dim Sql As String = "SELECT tax_code_id FROM ADDRESS WHERE cst_no=" & Me.Database.Escape(CustomerNo)
            Dim Id As Integer
            Try
                Id = Me.Database.GetOne(Sql)
                Me.Open(Id)
            Catch ex As Exception
                ' ignore
            End Try
        End If
    End Function

    Public Function CalculateTax(ByVal LineItems As LineItem(), Optional ByVal Freight As Double = 0) As Double
        If Me.Id > 0 Then
            Dim TotalTax As Double = 0
            Dim TaxAccounts As DataTable = Me.ExtractTaxAccounts(LineItems, Freight)
            For Each Row As DataRow In TaxAccounts.Rows
                TotalTax += Row.Item("amount")
            Next
            Return Math.Round(TotalTax, 2)
        Else
            Return 0
        End If
    End Function

    Public Function ExtractTaxAccounts(ByVal LineItems As LineItem(), Optional ByVal Freight As Double = 0) As DataTable
        Dim h As New Hashtable
        ' Create Tax Accounts TAble
        Dim TaxAccounts As New DataTable
        TaxAccounts.Columns.Add("account_no")
        TaxAccounts.Columns.Add("amount")
        TaxAccounts.Columns("amount").DefaultValue = 0
        ' Loop through each tax authority
        For Each a As DataRow In Me.TaxAuthorities.Rows
            ' Get account number for this tax authority
            Dim acctno As String = a.Item("gl_account")
            ' Create new tax account row if not yet added
            If Not h.ContainsKey(acctno) Then
                Dim r As DataRow = TaxAccounts.NewRow
                h.Add(acctno, TaxAccounts.Rows.Count)
                r.Item("account_no") = acctno
                r.Item("amount") = 0
                TaxAccounts.Rows.Add(r)
            End If
            ' Get Taxable amount
            Dim TaxSubTotal As Double = Me.GetTaxableAmount(a, LineItems, Freight)
            ' Add to appropriate account            
            TaxAccounts.Rows(h(acctno)).Item("amount") += Math.Round(TaxSubTotal, 2)
        Next
        Return TaxAccounts
    End Function

    Public Function GetTaxableAmount(ByVal a As DataRow, ByVal LineItems As LineItem(), Optional ByVal Freight As Double = 0) As Double
        ' Set subtotal at zero to start
        Dim TaxSubTotal As Double = 0
        ' Decide what to do based on authority
        If a.Item("formula_base") = 0 Then
            ' Flat Rate
            Dim Total As Double = 0
            For i As Integer = 0 To LineItems.Length - 1
                If LineItems(i).Taxable Then
                    Total += Math.Round(LineItems(i).Quantity * LineItems(i).Amount, 2)
                End If
            Next
            TaxSubTotal += Total * (a.Item("rate") / 100)
            ' Freight?
            If Me.FreightMethod = FreightTaxMethod.Always Or (Me.FreightMethod = FreightTaxMethod.WhenSomethingElseTaxed And Total > 0) Then
                TaxSubTotal += Freight * (a.Item("rate") / 100)
            End If
        ElseIf a.Item("formula_base") = 1 Then
            ' Calculate on total
            Dim Total As Double = 0
            For i As Integer = 0 To LineItems.Length - 1
                If LineItems(i).Taxable Then
                    Total += Math.Round(LineItems(i).Quantity * LineItems(i).Amount, 2)
                End If
            Next
            If Me.FreightMethod = FreightTaxMethod.Always Or (Me.FreightMethod = FreightTaxMethod.WhenSomethingElseTaxed And Total > 0) Then
                Total += Freight
            End If
            If Total > a.Item("dollar_limit") Then
                Dim OverLimit As Double = Total - a.Item("dollar_limit")
                Dim UnderLimit As Double = a.Item("dollar_limit")
                TaxSubTotal += UnderLimit * (a.Item("rate") / 100)
                TaxSubTotal += OverLimit * (a.Item("above_limit_rate") / 100)
            Else
                TaxSubTotal += Total * (a.Item("rate") / 100)
            End If
        Else
            ' Calculate per line item
            For i As Integer = 0 To LineItems.Length - 1
                If LineItems(i).Taxable Then
                    Dim Ext As Double = 0
                    If LineItems(i).Amount > a.Item("dollar_limit") Then
                        Ext += LineItems(i).Quantity * a.Item("dollar_limit") * (a.Item("rate") / 100)
                        Ext += (LineItems(i).Amount - a.Item("dollar_limit")) * LineItems(i).Quantity * (a.Item("above_limit_rate") / 100)
                    Else
                        Ext += LineItems(i).Quantity * LineItems(i).Amount * (a.Item("rate") / 100)
                    End If
                    TaxSubTotal += Math.Round(Ext, 2)
                End If
            Next
            ' Calculate tax on freight
            If Me.FreightMethod = FreightTaxMethod.Always Or (Me.FreightMethod = FreightTaxMethod.WhenSomethingElseTaxed And TaxSubTotal > 0) Then
                Dim Ext As Double = 0
                If Freight > a.Item("dollar_limit") Then
                    Ext += a.Item("dollar_limit") * (a.Item("rate") / 100)
                    Ext += (Freight - a.Item("dollar_limit")) * (a.Item("above_limit_rate") / 100)
                Else
                    Ext += Freight * (a.Item("rate") / 100)
                End If
                TaxSubTotal += Math.Round(Ext, 2)
            End If
        End If
        Return TaxSubTotal
    End Function

    Public Function TaxPerAuthority(ByVal LineItems As LineItem(), Optional ByVal Freight As Double = 0) As DataTable
        ' Create Table
        Dim Table As New DataTable
        Table.Columns.Add("id")
        Table.Columns.Add("name")
        Table.Columns.Add("rate").DefaultValue = 0
        Table.Columns.Add("amount").DefaultValue = 0
        ' Loop through each tax authority
        For Each a As DataRow In Me.TaxAuthorities.Rows
            ' Get Taxable amount
            Dim TaxSubTotal As Double = Me.GetTaxableAmount(a, LineItems, Freight)
            ' Add to table
            Dim r As DataRow = Table.NewRow
            r.Item("id") = a.Item("id")
            r.Item("name") = a.Item("name")
            r.Item("rate") = a.Item("rate")
            r.Item("amount") = TaxSubTotal
            Table.Rows.Add(r)
        Next
        Return Table
    End Function


End Class
