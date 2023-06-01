Imports MyCore.Data

Public Class cRateCalculator

    Public Database As EasySql

    ' Quantities
    Public StandardHours As Double = 0
    Public OvertimeHours As Double = 0
    Public HolidayHours As Double = 0
    Public TruckHours As Double = 0
    Public Mileage As Integer = 0
    Public NumTechs As Integer = 1
    Public NumHelpers As Integer = 0

    ' Properties
    Public RateTableId As Integer = 0
    Public Zone As Integer = 0
    Public TruckTypeId As Integer = 0
    Public RateTypeId As Integer = 0
    Public Contract As Boolean = False
    Public WorkTypeId As Integer = 0
    Public ChargeTax As Boolean = False
    Public ShopWork As Boolean = False
    Public Office As Integer = 0
    Public CustomerPickup As Boolean = False
    Public TaxCodeId As Integer = 0

    ' Line Items
    Public LineItems As DataTable

    Public ReadOnly Property LaborTotal() As Double
        Get
            Dim amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                If Row.Item("type_id") = 2 Then
                    amt += Row.Item("quantity") * Row.Item("unit")
                End If
            Next
            Return amt
        End Get
    End Property

    Public ReadOnly Property LaborUnit() As Double
        Get
            Dim amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                If Row.Item("type_id") = 2 Then
                    amt += Row.Item("unit")
                End If
            Next
            Return amt
        End Get
    End Property

    Public ReadOnly Property TravelTotal() As Double
        Get
            Dim amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                If Row.Item("type_id") = 3 Then
                    amt += Row.Item("quantity") * Row.Item("unit")
                End If
            Next
            Return amt
        End Get
    End Property

    Public ReadOnly Property TravelUnit() As Double
        Get
            Dim amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                If Row.Item("type_id") = 3 Then
                    amt += Row.Item("unit")
                End If
            Next
            Return amt
        End Get
    End Property

    Public ReadOnly Property Total() As Double
        Get
            Dim amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                amt += Row.Item("quantity") * Row.Item("unit")
            Next
            Return amt
        End Get
    End Property


    Public Sub New(ByVal db As EasySql)
        Me.Database = db
        Me.ResetLineItems()
    End Sub

    Private Sub ResetLineItems()
        Me.LineItems = New DataTable
        LineItems.Columns.Add("quantity", GetType(System.Double))
        LineItems.Columns.Add("part_no", GetType(System.String))
        LineItems.Columns.Add("description", GetType(System.String))
        LineItems.Columns.Add("type_id", GetType(System.String))
        LineItems.Columns.Add("unit", GetType(System.String))
        LineItems.Columns("type_id").DefaultValue = 0
    End Sub

    Public Function SetByCompany(ByVal CustomerNo As String) As Integer
        If CustomerNo.Length > 0 Then
            Dim Sql As String = "SELECT rate_table_id, tax_code_id, cst_zip, roundtrip_miles FROM ADDRESS WHERE cst_no=" & Me.Database.Escape(CustomerNo)
            Dim Row As DataRow = Me.Database.GetRow(Sql)
            If Me.Database.LastQuery.RowsReturned = 1 Then
                If Row.Item("tax_code_id") IsNot DBNull.Value Then
                    Me.TaxCodeId = Row.Item("tax_code_id")
                End If
                If Row.Item("rate_table_id") <> Me.RateTableId Then
                    Me.RateTableId = Row.Item("rate_table_id")
                End If
                If Row.Item("cst_zip") IsNot DBNull.Value Then
                    Me.Zone = Me.GetZoneFromZip(Row.Item("cst_zip"), 0)
                End If
                Me.Mileage = Row.Item("roundtrip_miles")
            End If
        End If
        Return Me.Zone
    End Function

    Public Function GetPrimaryOfficeFromZip(ByVal ZipCode As String) As Integer
        If ZipCode.Contains("-") Then
            ZipCode = ZipCode.Substring(0, ZipCode.IndexOf("-"))
        End If
        Dim off As Integer = 0
        Try
            off = Me.Database.GetOne("SELECT primary_office FROM zips WHERE zip=" & Me.Database.Escape(ZipCode))
        Catch ex As Exception
            Throw New Exception("Primary office is not set for zip code " & ZipCode)
        End Try
        Return off
    End Function

    Public Function GetZoneFromZip(ByVal ZipCode As String, Optional ByVal Office As Integer = 0) As Integer
        If ZipCode.Contains("-") Then
            ZipCode = ZipCode.Substring(0, ZipCode.IndexOf("-"))
        End If
        If Office = 0 Then
            Office = Me.GetPrimaryOfficeFromZip(ZipCode)
        End If
        Dim z As Integer = 0
        Try
            z = Me.Database.GetOne("SELECT zone FROM zip_to_zone WHERE zip=" & Me.Database.Escape(ZipCode) & " AND office=" & Office)
        Catch
            Throw New Exception("Zone is not set for zip code " & ZipCode & " and primary office " & Office)
        End Try
        Return z
    End Function

    Public Sub AutoAddLineItems(Optional ByVal AddLabor As Boolean = True, Optional ByVal AddBlanks As Boolean = False)
        Me.ResetLineItems()
        ' Get Rate Table
        Dim RateTable As New cRateTable(Me.Database)
        Try
            RateTable.Open(Me.RateTableId)
        Catch ex As Exception
            Throw New cException(cException.SeverityRating.Serious, "Not a valid rate table.", ex.ToString)
            Exit Sub
        End Try
        ' Mileage
        If RateTable.ChargeMileage Then
            Dim PartNo As String = Me.Database.GetOne("SELECT part_no FROM mileage_to_part_no WHERE truck_type_id=" & Me.TruckTypeId)
            If PartNo IsNot DBNull.Value Then
                Dim Sql As String = ""
                Sql &= "SELECT order_description, list_price, rti.price AS alternate_price, i.item_type_id"
                Sql &= " FROM item_master i"
                Sql &= " LEFT OUTER JOIN rate_table_item rti ON i.part_no=rti.part_no"
                Sql &= " AND rti.rate_table_id=" & Me.RateTableId
                Sql &= " WHERE i.part_no=" & Me.Database.Escape(PartNo)
                Dim Row As DataRow = Me.Database.GetRow(Sql)
                Dim NewRow As DataRow = Me.LineItems.NewRow
                NewRow.Item("quantity") = Me.Mileage
                NewRow.Item("part_no") = PartNo
                NewRow.Item("type_id") = Row.Item("item_type_id")
                NewRow.Item("description") = Row.Item("order_description")
                If Row.Item("alternate_price") Is DBNull.Value Then
                    NewRow.Item("unit") = Row.Item("list_price")
                Else
                    NewRow.Item("unit") = Row.Item("alternate_price")
                End If
                Me.LineItems.Rows.Add(NewRow)
            ElseIf AddBlanks Then
                Dim NewRow As DataRow = Me.LineItems.NewRow
                NewRow.Item("quantity") = Me.Mileage
                NewRow.Item("part_no") = ""
                NewRow.Item("description") = "Mileage"
                NewRow.Item("unit") = 0
                Me.LineItems.Rows.Add(NewRow)
            End If
        End If
        ' Truck Houly Rate
        If Me.TruckHours > 0 And RateTable.ChargeTruckHourly Then
            Dim PartNo As String = Me.Database.GetOne("SELECT hourly_part_no FROM truck_type WHERE id=" & Me.TruckTypeId)
            If PartNo IsNot DBNull.Value Then
                Dim Sql As String = ""
                Sql &= "SELECT order_description, list_price, rti.price AS alternate_price, i.item_type_id"
                Sql &= " FROM item_master i"
                Sql &= " LEFT OUTER JOIN rate_table_item rti ON i.part_no=rti.part_no"
                Sql &= " AND rti.rate_table_id=" & Me.RateTableId
                Sql &= " WHERE i.part_no=" & Me.Database.Escape(PartNo)
                Dim Row As DataRow = Me.Database.GetRow(Sql)
                Dim NewRow As DataRow = Me.LineItems.NewRow
                NewRow.Item("quantity") = Me.TruckHours
                NewRow.Item("part_no") = PartNo
                NewRow.Item("type_id") = Row.Item("item_type_id")
                NewRow.Item("description") = Row.Item("order_description")
                If Row.Item("alternate_price") Is DBNull.Value Then
                    NewRow.Item("unit") = Row.Item("list_price")
                Else
                    NewRow.Item("unit") = Row.Item("alternate_price")
                End If
                Me.LineItems.Rows.Add(NewRow)
            ElseIf AddBlanks Then
                Dim NewRow As DataRow = Me.LineItems.NewRow
                NewRow.Item("quantity") = Me.TruckHours
                NewRow.Item("part_no") = ""
                NewRow.Item("description") = "Truck Hours"
                NewRow.Item("unit") = 0
                Me.LineItems.Rows.Add(NewRow)
            End If
        End If
        ' Labor Hourly Rate
        If AddLabor Then
            ' Get Hours
            Dim Std As Double = Me.StandardHours
            Dim Ot As Double = Me.OvertimeHours
            Dim Hol As Double = Me.HolidayHours
            ' Adjust hours if not over minimum
            If Std + Ot + Hol < RateTable.MinimumHours Then
                Dim Diff As Double = RateTable.MinimumHours - Std - Ot - Hol
                If Std > 0 Then
                    Std += Diff
                ElseIf Ot > 0 Then
                    Ot += Diff
                ElseIf Hol > 0 Then
                    Hol += Diff
                Else
                    Std += Diff
                End If
            End If
            ' Set rate
            If Std > 0 Or Ot > 0 Or Hol > 0 Then
                Dim Sql As String = ""
                Sql &= "SELECT standard_part_no, overtime_part_no, holiday_part_no FROM hourly_to_part_no"
                Sql &= " WHERE work_type_id=" & Me.WorkTypeId
                If RateTable.ContractDiscount Then
                    Sql &= " AND contract=" & Me.Database.Escape(Me.Contract)
                Else
                    Sql &= " AND contract=0"
                End If
                Dim pn As DataRow = Me.Database.GetRow(Sql)
                If Std > 0 Then
                    If pn.Item("standard_part_no") IsNot DBNull.Value Then
                        Sql = "SELECT order_description, list_price, rti.price AS alternate_price, i.item_type_id"
                        Sql &= " FROM item_master i"
                        Sql &= " LEFT OUTER JOIN rate_table_item rti ON i.part_no=rti.part_no"
                        Sql &= " AND rti.rate_table_id=" & Me.RateTableId
                        Sql &= " WHERE i.part_no=" & Me.Database.Escape(pn.Item("standard_part_no"))
                        Dim Row As DataRow = Me.Database.GetRow(Sql)
                        If Me.Database.LastQuery.RowsReturned = 1 Then
                            Dim NewRow As DataRow = Me.LineItems.NewRow
                            NewRow.Item("quantity") = Std * Me.NumTechs
                            NewRow.Item("part_no") = pn.Item("standard_part_no")
                            NewRow.Item("type_id") = Row.Item("item_type_id")
                            NewRow.Item("description") = Row.Item("order_description")
                            If Row.Item("alternate_price") Is DBNull.Value Then
                                NewRow.Item("unit") = Row.Item("list_price")
                            Else
                                NewRow.Item("unit") = Row.Item("alternate_price")
                            End If
                            Me.LineItems.Rows.Add(NewRow)
                        Else
                            Dim NewRow As DataRow = Me.LineItems.NewRow
                            NewRow.Item("quantity") = Std * Me.NumTechs
                            NewRow.Item("part_no") = pn.Item("standard_part_no")
                            NewRow.Item("description") = "Part No Not Found"
                            NewRow.Item("unit") = 0
                            Me.LineItems.Rows.Add(NewRow)
                        End If
                    ElseIf AddBlanks Then
                        Dim NewRow As DataRow = Me.LineItems.NewRow
                        NewRow.Item("quantity") = Std * Me.NumTechs
                        NewRow.Item("part_no") = ""
                        NewRow.Item("description") = "Standard Labor"
                        NewRow.Item("unit") = 0
                        Me.LineItems.Rows.Add(NewRow)
                    End If
                End If
                If Ot > 0 Then
                    If pn.Item("overtime_part_no") IsNot DBNull.Value Then
                        Sql = "SELECT order_description, list_price, rti.price AS alternate_price, i.item_type_id"
                        Sql &= " FROM item_master i"
                        Sql &= " LEFT OUTER JOIN rate_table_item rti ON i.part_no=rti.part_no"
                        Sql &= " AND rti.rate_table_id=" & Me.RateTableId
                        Sql &= " WHERE i.part_no=" & Me.Database.Escape(pn.Item("overtime_part_no"))
                        Dim Row As DataRow = Me.Database.GetRow(Sql)
                        If Me.Database.LastQuery.RowsReturned = 1 Then
                            Dim NewRow As DataRow = Me.LineItems.NewRow
                            NewRow.Item("quantity") = Ot * Me.NumTechs
                            NewRow.Item("part_no") = pn.Item("overtime_part_no")
                            NewRow.Item("type_id") = Row.Item("item_type_id")
                            NewRow.Item("description") = Row.Item("order_description")
                            If Row.Item("alternate_price") Is DBNull.Value Then
                                NewRow.Item("unit") = Row.Item("list_price")
                            Else
                                NewRow.Item("unit") = Row.Item("alternate_price")
                            End If
                            Me.LineItems.Rows.Add(NewRow)
                        Else
                            Dim NewRow As DataRow = Me.LineItems.NewRow
                            NewRow.Item("quantity") = Ot * Me.NumTechs
                            NewRow.Item("part_no") = pn.Item("overtimepart_no")
                            NewRow.Item("description") = "Part No Not Found"
                            NewRow.Item("unit") = 0
                            Me.LineItems.Rows.Add(NewRow)
                        End If
                    ElseIf AddBlanks Then
                        Dim NewRow As DataRow = Me.LineItems.NewRow
                        NewRow.Item("quantity") = Ot * Me.NumTechs
                        NewRow.Item("part_no") = ""
                        NewRow.Item("description") = "Overtime Labor"
                        NewRow.Item("unit") = 0
                        Me.LineItems.Rows.Add(NewRow)
                    End If
                End If
                If Hol > 0 Then
                    If pn.Item("holiday_part_no") IsNot DBNull.Value Then
                        Sql = "SELECT order_description, list_price, rti.price AS alternate_price, i.item_type_id"
                        Sql &= " FROM item_master i"
                        Sql &= " LEFT OUTER JOIN rate_table_item rti ON i.part_no=rti.part_no"
                        Sql &= " AND rti.rate_table_id=" & Me.RateTableId
                        Sql &= " WHERE i.part_no=" & Me.Database.Escape(pn.Item("holiday_part_no"))
                        Dim Row As DataRow = Me.Database.GetRow(Sql)
                        If Me.Database.LastQuery.RowsReturned = 1 Then
                            Dim NewRow As DataRow = Me.LineItems.NewRow
                            NewRow.Item("quantity") = Hol * Me.NumTechs
                            NewRow.Item("part_no") = pn.Item("holiday_part_no")
                            NewRow.Item("type_id") = Row.Item("item_type_id")
                            NewRow.Item("description") = Row.Item("order_description")
                            If Row.Item("alternate_price") Is DBNull.Value Then
                                NewRow.Item("unit") = Row.Item("list_price")
                            Else
                                NewRow.Item("unit") = Row.Item("alternate_price")
                            End If
                            Me.LineItems.Rows.Add(NewRow)
                        Else
                            Dim NewRow As DataRow = Me.LineItems.NewRow
                            NewRow.Item("quantity") = Hol * Me.NumTechs
                            NewRow.Item("part_no") = pn.Item("holiday_part_no")
                            NewRow.Item("description") = "Part No Not Found"
                            NewRow.Item("unit") = 0
                            Me.LineItems.Rows.Add(NewRow)
                        End If
                    ElseIf AddBlanks Then
                        Dim NewRow As DataRow = Me.LineItems.NewRow
                        NewRow.Item("quantity") = Hol * Me.NumTechs
                        NewRow.Item("part_no") = ""
                        NewRow.Item("description") = "Holiday Labor"
                        NewRow.Item("unit") = 0
                        Me.LineItems.Rows.Add(NewRow)
                    End If
                End If
            End If
        End If
        ' Zone Charges
        If Me.Zone > 0 And RateTable.ChargeZone Then
            Dim Sql As String = ""
            Sql &= "SELECT truck_part_no, tech_part_no FROM zone_to_part_no"
            Sql &= " WHERE zone_id=" & Me.Zone
            Sql &= " AND truck_type_id=" & Me.TruckTypeId
            Sql &= " AND rate_type_id=" & Me.RateTypeId
            Dim pn As DataRow = Me.Database.GetRow(Sql)
            If Me.Database.LastQuery.RowsReturned = 1 Then
                If pn.Item("truck_part_no") IsNot DBNull.Value Then
                    Sql = "SELECT order_description, list_price, rti.price AS alternate_price, i.item_type_id"
                    Sql &= " FROM item_master i"
                    Sql &= " LEFT OUTER JOIN rate_table_item rti ON i.part_no=rti.part_no"
                    Sql &= " AND rti.rate_table_id=" & Me.RateTableId
                    Sql &= " WHERE i.part_no=" & Me.Database.Escape(pn.Item("truck_part_no"))
                    Dim Row As DataRow = Me.Database.GetRow(Sql)
                    If Me.Database.LastQuery.RowsReturned = 1 Then
                        Dim NewRow As DataRow = Me.LineItems.NewRow
                        NewRow.Item("quantity") = 1
                        NewRow.Item("part_no") = pn.Item("truck_part_no")
                        NewRow.Item("type_id") = Row.Item("item_type_id")
                        NewRow.Item("description") = Row.Item("order_description")
                        If Row.Item("alternate_price") Is DBNull.Value Then
                            NewRow.Item("unit") = Row.Item("list_price")
                        Else
                            NewRow.Item("unit") = Row.Item("alternate_price")
                        End If
                        Me.LineItems.Rows.Add(NewRow)
                    Else
                        Dim NewRow As DataRow = Me.LineItems.NewRow
                        NewRow.Item("quantity") = 1
                        NewRow.Item("part_no") = pn.Item("truck_part_no")
                        NewRow.Item("description") = "Part No Not Found"
                        NewRow.Item("unit") = 0
                        Me.LineItems.Rows.Add(NewRow)
                    End If
                ElseIf AddBlanks Then
                    Dim NewRow As DataRow = Me.LineItems.NewRow
                    NewRow.Item("quantity") = 1
                    NewRow.Item("part_no") = ""
                    NewRow.Item("description") = "Truck Zone"
                    NewRow.Item("unit") = 0
                    Me.LineItems.Rows.Add(NewRow)
                End If
                If pn.Item("tech_part_no") IsNot DBNull.Value Then
                    Sql = "SELECT order_description, list_price, rti.price AS alternate_price, i.item_type_id"
                    Sql &= " FROM item_master i"
                    Sql &= " LEFT OUTER JOIN rate_table_item rti ON i.part_no=rti.part_no"
                    Sql &= " AND rti.rate_table_id=" & Me.RateTableId
                    Sql &= " WHERE i.part_no=" & Me.Database.Escape(pn.Item("tech_part_no"))
                    Dim Row As DataRow = Me.Database.GetRow(Sql)
                    If Me.Database.LastQuery.RowsReturned = 1 Then
                        Dim NewRow As DataRow = Me.LineItems.NewRow
                        NewRow.Item("quantity") = Me.NumTechs + Me.NumHelpers
                        NewRow.Item("part_no") = pn.Item("tech_part_no")
                        NewRow.Item("type_id") = Row.Item("item_type_id")
                        NewRow.Item("description") = Row.Item("order_description")
                        If Row.Item("alternate_price") Is DBNull.Value Then
                            NewRow.Item("unit") = Row.Item("list_price")
                        Else
                            NewRow.Item("unit") = Row.Item("alternate_price")
                        End If
                        Me.LineItems.Rows.Add(NewRow)
                    Else
                        Dim NewRow As DataRow = Me.LineItems.NewRow
                        NewRow.Item("quantity") = Me.NumTechs + Me.NumHelpers
                        NewRow.Item("part_no") = pn.Item("tech_part_no")
                        NewRow.Item("description") = "Part No Not Found"
                        NewRow.Item("unit") = 0
                        Me.LineItems.Rows.Add(NewRow)
                    End If
                ElseIf AddBlanks Then
                    Dim NewRow As DataRow = Me.LineItems.NewRow
                    NewRow.Item("quantity") = Me.NumTechs + Me.NumHelpers
                    NewRow.Item("part_no") = ""
                    NewRow.Item("description") = "Tech Zone"
                    NewRow.Item("unit") = 0
                    Me.LineItems.Rows.Add(NewRow)
                End If
            ElseIf AddBlanks Then
                ' No part number defined for this
                ' Truck zone
                Dim NewRow As DataRow = Me.LineItems.NewRow
                NewRow.Item("quantity") = 1
                NewRow.Item("part_no") = ""
                NewRow.Item("description") = "Truck Zone"
                NewRow.Item("unit") = 0
                Me.LineItems.Rows.Add(NewRow)
                ' Tech Zone
                NewRow = Me.LineItems.NewRow
                NewRow.Item("quantity") = Me.NumTechs + Me.NumHelpers
                NewRow.Item("part_no") = ""
                NewRow.Item("description") = "Tech Zone"
                NewRow.Item("unit") = 0
                Me.LineItems.Rows.Add(NewRow)
            End If
        End If
        ' Enforce valid values in line items
        For Each Row As DataRow In Me.LineItems.Rows
            ' Quantity must be at least one
            If Row.Item("quantity") Is DBNull.Value Then
                Row.Item("quantity") = 1
            ElseIf Row.Item("quantity") = Nothing Then
                Row.Item("quantity") = 1
            ElseIf Not Microsoft.VisualBasic.IsNumeric(Row.Item("quantity")) Then
                Row.Item("quantity") = 1
            ElseIf Row.Item("quantity") <= 0 Then
                Row.Item("quantity") = 1
            End If
            ' Unit must be at least one
            If Row.Item("unit") Is DBNull.Value Then
                Row.Item("unit") = 0
            ElseIf Row.Item("quantity") = Nothing Then
                Row.Item("unit") = 0
            ElseIf Not Microsoft.VisualBasic.IsNumeric(Row.Item("unit")) Then
                Row.Item("unit") = 0
            End If
        Next
    End Sub

    Public Function ZoneCharge() As Double
        Dim RateTable As New cRateTable(Me.Database)
        Dim Amount As Double = 0
        If Me.Zone > 0 And RateTable.ChargeZone Then
            Dim Sql As String = ""
            Sql &= "SELECT truck_part_no, tech_part_no FROM zone_to_part_no"
            Sql &= " WHERE zone_id=" & Me.Zone
            Sql &= " AND truck_type_id=" & Me.TruckTypeId
            Sql &= " AND rate_type_id=" & Me.RateTypeId
            Dim pn As DataRow = Me.Database.GetRow(Sql)
            If Me.Database.LastQuery.RowsReturned = 1 Then
                If pn.Item("truck_part_no") IsNot DBNull.Value Then
                    Sql = "SELECT list_price, rti.price AS alternate_price"
                    Sql &= " FROM item_master i"
                    Sql &= " LEFT OUTER JOIN rate_table_item rti ON i.part_no=rti.part_no"
                    Sql &= " AND rti.rate_table_id=" & Me.RateTableId
                    Sql &= " WHERE i.part_no=" & Me.Database.Escape(pn.Item("truck_part_no"))
                    Dim Row As DataRow = Me.Database.GetRow(Sql)
                    If Me.Database.LastQuery.RowsReturned = 1 Then
                        If Row.Item("alternate_price") Is DBNull.Value Then
                            Amount += Row.Item("list_price")
                        Else
                            Amount += Row.Item("alternate_price")
                        End If
                    End If
                End If
                If pn.Item("tech_part_no") IsNot DBNull.Value Then
                    Sql = "SELECT list_price, rti.price AS alternate_price"
                    Sql &= " FROM item_master i"
                    Sql &= " LEFT OUTER JOIN rate_table_item rti ON i.part_no=rti.part_no"
                    Sql &= " AND rti.rate_table_id=" & Me.RateTableId
                    Sql &= " WHERE i.part_no=" & Me.Database.Escape(pn.Item("tech_part_no"))
                    Dim Row As DataRow = Me.Database.GetRow(Sql)
                    If Me.Database.LastQuery.RowsReturned = 1 Then
                        If Row.Item("alternate_price") Is DBNull.Value Then
                            Amount += Row.Item("list_price") * (Me.NumTechs + Me.NumHelpers)
                        Else
                            Amount += Row.Item("alternate_price") * (Me.NumTechs + Me.NumHelpers)
                        End If
                    End If
                End If
            End If
        End If
        Return Amount
    End Function

    Public Function GetLabor() As Double
        Dim RateTable As New cRateTable(Me.Database)
        Dim Amount As Double = 0
        If Me.StandardHours > 0 Or Me.OvertimeHours > 0 Or Me.HolidayHours > 0 Then
            Dim Sql As String = ""
            Sql &= "SELECT standard_part_no, overtime_part_no, holiday_part_no FROM hourly_to_part_no"
            Sql &= " WHERE work_type_id=" & Me.WorkTypeId
            If RateTable.ContractDiscount Then
                Sql &= " AND contract=" & Me.Database.Escape(Me.Contract)
            Else
                Sql &= " AND contract=0"
            End If
            Dim pn As DataRow = Me.Database.GetRow(Sql)
            If Me.StandardHours > 0 Then
                If pn.Item("standard_part_no") IsNot DBNull.Value Then
                    Sql = "SELECT list_price, rti.price AS alternate_price"
                    Sql &= " FROM item_master i"
                    Sql &= " LEFT OUTER JOIN rate_table_item rti ON i.part_no=rti.part_no"
                    Sql &= " AND rti.rate_table_id=" & Me.RateTableId
                    Sql &= " WHERE i.part_no=" & Me.Database.Escape(pn.Item("standard_part_no"))
                    Dim Row As DataRow = Me.Database.GetRow(Sql)
                    If Me.Database.LastQuery.RowsReturned = 1 Then
                        If Row.Item("alternate_price") Is DBNull.Value Then
                            Amount += Row.Item("list_price") * (Me.NumTechs + Me.NumHelpers)
                        Else
                            Amount += Row.Item("alternate_price") * (Me.NumTechs + Me.NumHelpers)
                        End If
                    End If
                End If
            End If
            If Me.OvertimeHours > 0 Then
                If pn.Item("overtime_part_no") IsNot DBNull.Value Then
                    Sql = "SELECT list_price, rti.price AS alternate_price"
                    Sql &= " FROM item_master i"
                    Sql &= " LEFT OUTER JOIN rate_table_item rti ON i.part_no=rti.part_no"
                    Sql &= " AND rti.rate_table_id=" & Me.RateTableId
                    Sql &= " WHERE i.part_no=" & Me.Database.Escape(pn.Item("overtime_part_no"))
                    Dim Row As DataRow = Me.Database.GetRow(Sql)
                    If Me.Database.LastQuery.RowsReturned = 1 Then
                        If Row.Item("alternate_price") Is DBNull.Value Then
                            Amount += Row.Item("list_price") * (Me.NumTechs + Me.NumHelpers)
                        Else
                            Amount += Row.Item("alternate_price") * (Me.NumTechs + Me.NumHelpers)
                        End If
                    End If
                End If
            End If
            If Me.HolidayHours > 0 Then
                If pn.Item("holiday_part_no") IsNot DBNull.Value Then
                    Sql = "SELECT list_price, rti.price AS alternate_price"
                    Sql &= " FROM item_master i"
                    Sql &= " LEFT OUTER JOIN rate_table_item rti ON i.part_no=rti.part_no"
                    Sql &= " AND rti.rate_table_id=" & Me.RateTableId
                    Sql &= " WHERE i.part_no=" & Me.Database.Escape(pn.Item("holiday_part_no"))
                    Dim Row As DataRow = Me.Database.GetRow(Sql)
                    If Me.Database.LastQuery.RowsReturned = 1 Then
                        Dim NewRow As DataRow = Me.LineItems.NewRow
                        If Row.Item("alternate_price") Is DBNull.Value Then
                            Amount += Row.Item("list_price") * (Me.NumTechs + Me.NumHelpers)
                        Else
                            Amount += Row.Item("alternate_price") * (Me.NumTechs + Me.NumHelpers)
                        End If
                    End If
                End If
            End If
        End If
        Return Amount
    End Function



End Class
