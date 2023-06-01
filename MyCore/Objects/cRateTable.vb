Public Class cRateTable

    Dim _Id As Integer = 0
    Dim Database As MyCore.Data.EasySql

    Public Name As String = ""
    Public ChargeMileage As Boolean = False
    Public ChargeZone As Boolean = True
    Public ChargeTruckHourly As Boolean = False
    Public ContractDiscount As Boolean = True
    Public LaborChargeTimeBy As LaborMethod = LaborMethod.OnSite
    Public DefaultLateChargePercent As Double = 2
    Public HelperPercent As Double = 60
    Public MinimumHours As Double = 1
    Public BillingUnits As Integer = 15
    Public SalesPriceFormula As String = ""

    Public LastUpdatedBy As String = ""
    Public CreatedBy As String = ""

    Public PartNoPricing As DataTable

    Public Event Refresh()

    Public ReadOnly Property Id() As Integer
        Get
            Return Me._Id
        End Get
    End Property


    Public Enum LaborMethod
        OnSite = 0
        PortalToPortal = 1
        PortalToPortalActual = 2
    End Enum

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
    End Sub

    Public Sub OpenByCompany(ByVal CustomerNo As String)
        Dim rt As Integer = Me.Database.GetOne("SELECT rate_table_id FROM ADDRESS WHERE cst_no=" & Me.Database.Escape(CustomerNo))
        Me.Open(rt)
    End Sub

    Public Sub Open(ByVal Id As Integer)
        If Id > 0 Then
            Me._Id = Id
            Dim Row As DataRow = Me.Database.GetRow("SELECT * FROM rate_table WHERE id=" & Me._Id)
            If Me.Database.LastQuery.RowsReturned = 1 Then
                If Row IsNot Nothing Then
                    Me.Name = Row.Item("name")
                    Me.ChargeZone = Row.Item("zone_charge")
                    Me.ChargeTruckHourly = Row.Item("truck_hourly")
                    Me.ChargeMileage = Row.Item("mileage_charge")
                    Me.ContractDiscount = Row.Item("contract_rate")
                    Me.DefaultLateChargePercent = Row.Item("late_charge_percent")
                    Me.HelperPercent = Row.Item("helper_percent")
                    Me.MinimumHours = Row.Item("minimum_hours")
                    Me.BillingUnits = Row.Item("billing_units")
                    Me.SalesPriceFormula = Row.Item("sales_price_formula")
                    Me.CreatedBy = Row.Item("created_by")
                    Me.LastUpdatedBy = Row.Item("last_updated_by")
                    Select Case Row.Item("portal_to_portal")
                        Case 1
                            Me.LaborChargeTimeBy = LaborMethod.PortalToPortal
                        Case 2
                            Me.LaborChargeTimeBy = LaborMethod.PortalToPortalActual
                        Case Else
                            Me.LaborChargeTimeBy = LaborMethod.OnSite
                    End Select
                    Me.PartNoPricing = Me.Database.GetAll("SELECT * FROM rate_table_item WHERE rate_table_id=" & Me._Id)
                    RaiseEvent Refresh()
                Else
                    Me._Id = 0
                    Throw New Exception("Rate table not found.")
                End If
            Else
                Me._Id = 0
                Throw New Exception("Rate table not found.")
            End If
        End If
    End Sub

    Public Sub Save()
        If Me._Id = 0 Then
            ' Add new
            Dim Sql As String = ""
            Sql &= "INSERT INTO rate_table (name, zone_charge, mileage_charge, truck_hourly,"
            Sql &= " helper_percent, minimum_hours, billing_units, sales_price_formula,"
            Sql &= " portal_to_portal, contract_rate, late_charge_percent, "
            Sql &= " date_last_updated, date_created, last_updated_by, created_by"
            Sql &= ") VALUES (@name, @zone_charge, @mileage_charge, @truck_hourly,"
            Sql &= " @helper_percent, @minimum_hours, @billing_units, @sales_price_formula,"
            Sql &= " @portal_to_portal, @contract_rate, @late_charge_percent, "
            Sql &= " @now, @now, @user, @user)"
            Sql = Sql.Replace("@name", Me.Database.Escape(Me.Name))
            Sql = Sql.Replace("@zone_charge", IIf(Me.ChargeZone, 1, 0))
            Sql = Sql.Replace("@mileage_charge", IIf(Me.ChargeMileage, 1, 0))
            Sql = Sql.Replace("@truck_hourly", IIf(Me.ChargeTruckHourly, 1, 0))
            Select Case Me.LaborChargeTimeBy
                Case LaborMethod.OnSite
                    Sql = Sql.Replace("@portal_to_portal", "0")
                Case LaborMethod.PortalToPortal
                    Sql = Sql.Replace("@portal_to_portal", "1")
                Case LaborMethod.PortalToPortalActual
                    Sql = Sql.Replace("@portal_to_portal", "2")
            End Select
            Sql = Sql.Replace("@contract_rate", IIf(Me.ContractDiscount, 1, 0))
            Sql = Sql.Replace("@late_charge_percent", Me.Database.Escape(Me.DefaultLateChargePercent))
            Sql = Sql.Replace("@helper_percent", Me.Database.Escape(Me.HelperPercent))
            Sql = Sql.Replace("@minimum_hours", Me.Database.Escape(Me.MinimumHours))
            Sql = Sql.Replace("@billing_units", Me.Database.Escape(Me.BillingUnits))
            Sql = Sql.Replace("@sales_price_formula", Me.Database.Escape(Me.SalesPriceFormula))
            Sql = Sql.Replace("@now", Me.Database.Escape(Now))
            Sql = Sql.Replace("@user", Me.Database.Escape(Me.LastUpdatedBy))
            Me.Database.InsertAndReturnId(Sql)
        Else
            ' Edit
            Dim Sql As String = ""
            Sql &= "UPDATE rate_table SET "
            Sql &= "name=" & Me.Database.Escape(Me.Name) & ", "
            Sql &= "zone_charge=" & IIf(Me.ChargeZone, 1, 0) & ", "
            Sql &= "mileage_charge=" & IIf(Me.ChargeMileage, 1, 0) & ", "
            Sql &= "truck_hourly=" & IIf(Me.ChargeTruckHourly, 1, 0) & ", "
            Select Case Me.LaborChargeTimeBy
                Case LaborMethod.OnSite
                    Sql &= "portal_to_portal=0, "
                Case LaborMethod.PortalToPortal
                    Sql &= "portal_to_portal=1, "
                Case LaborMethod.PortalToPortalActual
                    Sql &= "portal_to_portal=2, "
            End Select
            Sql &= "contract_rate=" & IIf(Me.ContractDiscount, 1, 0) & ", "
            Sql &= "late_charge_percent=" & Me.Database.Escape(Me.DefaultLateChargePercent) & ", "
            Sql &= "helper_percent=" & Me.Database.Escape(Me.HelperPercent) & ", "
            Sql &= "minimum_hours=" & Me.MinimumHours & ", "
            Sql &= "billing_units=" & Me.BillingUnits & ", "
            Sql &= "sales_price_formula=" & Me.Database.Escape(Me.SalesPriceFormula) & ", "
            Sql &= "date_last_updated=" & Me.Database.Escape(Now) & ", "
            Sql &= "last_updated_by=" & Me.Database.Escape(Me.LastUpdatedBy)
            Sql &= " WHERE id=" & Me._Id
            Me.Database.Execute(Sql)
        End If
        If Me.Database.LastQuery.Successful Then
            If Me._Id = 0 Then
                Me._Id = Me.Database.LastQuery.InsertId
            End If
            Me.SavePartNoPricing()
            Me.Open(Me._Id)
        Else
            Throw New Exception("Save of rate table failed with error: " & Me.Database.LastQuery.ErrorMsg)
        End If
    End Sub

    Private Sub SavePartNoPricing()
        If Me.PartNoPricing IsNot Nothing Then
            Dim Sql As String = ""
            For Each Row As DataRow In Me.PartNoPricing.Rows
                If Row.RowState = DataRowState.Added Then
                    Sql = "INSERT INTO rate_table_item (rate_table_id, part_no, price,"
                    Sql &= " date_last_updated, last_updated_by)"
                    Sql &= " VALUES ("
                    Sql &= Me.Id & ", "
                    Sql &= Me.Database.Escape(Row.Item("part_no")) & ", "
                    Sql &= Me.Database.Escape(Row.Item("price")) & ", "
                    Sql &= Me.Database.Escape(Now) & ", "
                    Sql &= Me.Database.Escape(Me.LastUpdatedBy)
                    Sql &= ")"
                    Me.Database.Execute(Sql)
                    If Not Me.Database.LastQuery.Successful Then
                        Dim Err As String = Me.Database.LastQuery.ErrorMsg
                    End If
                ElseIf Row.RowState = DataRowState.Modified Then
                    Sql = "UPDATE rate_table_item SET "
                    Sql &= " part_no=" & Me.Database.Escape(Row.Item("part_no")) & ", "
                    Sql &= " price=" & Me.Database.Escape(Row.Item("price")) & ", "
                    Sql &= " date_last_updated=" & Me.Database.Escape(Now) & ", "
                    Sql &= " last_updated_by=" & Me.Database.Escape(Me.LastUpdatedBy)
                    Sql &= " WHERE id=" & Row.Item("id")
                    Me.Database.Execute(Sql)
                End If
            Next
        End If
    End Sub

    Public Function GetPrice(ByVal PartNo As String, Optional ByVal AllOrNothing As Boolean = False) As Double
        Dim Sql As String = ""
        Sql = "SELECT im.list_price, rti.price"
        Sql &= " FROM item_master im"
        Sql &= " LEFT OUTER JOIN rate_table_item rti ON im.part_no=rti.part_no AND rate_table_id=" & Me.Id
        Sql &= " WHERE im.part_no=" & Me.Database.Escape(PartNo)
        Dim Row As DataRow = Me.Database.GetRow(Sql)
        If Me.Database.LastQuery.RowsReturned = 1 Then
            If Row.Item("price") IsNot DBNull.Value Then
                Return Row.Item("price")
            Else
                If AllOrNothing Then
                    Return Nothing
                Else
                    Return Row.Item("list_price")
                End If
            End If
        Else
            Throw New Exception("Part number not found.")
        End If
    End Function


End Class
