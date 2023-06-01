Public Class cCalAgreement

    Dim _Id As Integer = Nothing
    Dim Database As MyCore.Data.EasySql
    Dim Settings As cSettings

    Public ShipTo As String = ""
    Public BillTo As String = ""
    Public Contact As String = ""
    Public ContactID As Integer = 0
    Public ServiceAreaId As Integer = 0
    Public TruckTypeId As Integer = 0
    Public FrequencyMonths As Integer = 3
    Public FrequencyDays As Integer = 0
    Public DateStart As DateTime = Now
    Public DateToReview As DateTime = Now.AddYears(1)
    Public DateNextCal As DateTime = Now
    Public Canceled As Boolean = False
    Public AppointmentOnly As Boolean = False
    Public StrictDeadline As Boolean = False
    Public NonRecurring As Boolean = False
    Public EstimatedHours As Double = 0
    Public InternalNotes As String = ""
    Public ExternalNotes As String = ""
    Public Po As String = ""
    Public Formula As String = ""
    Public LastUpdatedBy As String = ""
    Public CreatedBy As String = ""

    Public DateCancelled As DateTime
    Public CancelReason As String = ""
    Public CancelledBy As String = ""

    Public Equipment As DataTable
    Public LineItems As DataTable

    Public Event Reload()
    Public Event Saved(ByVal Agreement As cCalAgreement)

    Public ReadOnly Property Price() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.LineItems.Rows
                If Row.Item("quantity") Is DBNull.Value Then
                    Row.Item("quantity") = 1
                End If
                Amt += Row.Item("quantity") * Row.Item("price")
            Next
            Return Amt
        End Get
    End Property

    Public ReadOnly Property DatePreviousCal() As DateTime
        Get
            Dim Sql As String = "SELECT TOP 1 date_completed FROM service_order WHERE cal_agreement_id=" & Me.Id
            Sql &= " AND date_completed IS NOT NULL ORDER BY date_completed DESC"
            Dim Row As DataRow = Me.Database.GetRow(Sql)
            If Me.Database.LastQuery.RowsReturned = 1 Then
                Return Row.Item("date_completed")
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property DatePreviousCalDue() As DateTime
        Get
            Dim Sql As String = "SELECT TOP 1 date_due FROM service_order WHERE cal_agreement_id=" & Me.Id
            Sql &= " ORDER BY date_completed DESC"
            Dim Row As DataRow = Me.Database.GetRow(Sql)
            If Me.Database.LastQuery.RowsReturned = 1 Then
                Return Row.Item("date_due")
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property Id() As Integer
        Get
            Return Me._Id
        End Get
    End Property

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
        Me.Settings = New cSettings(Me.Database)
    End Sub

    Public Sub OpenAsNew()
        Me.CreateBlankEquipmentTable()
        ' Blank line items
        Me.LineItems = New DataTable
        Me.LineItems.Columns.Add("id")
        Me.LineItems.Columns.Add("quantity")
        Me.LineItems.Columns.Add("part_no")
        Me.LineItems.Columns.Add("description")
        Me.LineItems.Columns.Add("price")
        Me.LineItems.Columns.Add("tax_status_id")
        ' event
        RaiseEvent Reload()
    End Sub

    Public Sub Open(ByVal Id As Integer)
        ' Get Agreement Data
        Dim Sql As String = "SELECT cal.* "
        Sql &= " FROM cal_agreement cal"
        Sql &= " WHERE id = " & Id
        Dim Row As DataRow = Me.Database.GetRow(Sql)
        If Me.Database.LastQuery.Successful Then
            Me._Id = Row.Item("id")
            Me.DateStart = Row.Item("date_start")
            Me.DateToReview = Row.Item("date_end")
            Me.Po = Me.IsNull(Row.Item("purchase_order"), "")
            Me.ExternalNotes = Me.IsNull(Row.Item("comments"), "")
            Me.InternalNotes = Me.IsNull(Row.Item("internal_notes"), "")
            Me.ShipTo = Me.IsNull(Row.Item("ship_to_no"), "")
            Me.BillTo = Me.IsNull(Row.Item("bill_to_no"), "")
            Me.Contact = Me.IsNull(Row.Item("contact_name"), "")
            Me.ContactID = Row.Item("contact_id")
            Me.FrequencyMonths = Row.Item("frequency_months")
            Me.FrequencyDays = Row.Item("frequency_days")
            Me.StrictDeadline = Row.Item("strict_deadline")
            Me.AppointmentOnly = Row.Item("appointment_only")
            Me.ServiceAreaId = Row.Item("service_area_id")
            Me.EstimatedHours = Row.Item("estimated_hours")
            Me.Canceled = Row.Item("canceled")
            Me.Formula = Row.Item("formula")
            Me.NonRecurring = Row.Item("non_recurring")
            Me.TruckTypeId = Row.Item("truck_type_id")
            Me.CreatedBy = Me.IsNull(Row.Item("created_by"), "")
            Me.LastUpdatedBy = Row.Item("last_updated_by")
            Me.DateCancelled = IIf(Row.Item("date_cancel") Is DBNull.Value, Nothing, Row.Item("date_cancel"))
            Me.CancelReason = Me.IsNull(Row.Item("cancel_reason"), "")
            Me.CancelledBy = Me.IsNull(Row.Item("cancelled_by"), "")
            Me.DateNextCal = IIf(Row.Item("date_next") Is DBNull.Value, Nothing, Row.Item("date_next"))
            ' Get Line Items
            Me.LineItems = Me.Database.GetAll("SELECT id, quantity, part_no, description, price, tax_status_id FROM cal_agreement_item WHERE cal_agreement_id=" & Me.Id)
            ' Get Equipment
            Sql = "SELECT DISTINCT dep_id AS id,"
            Sql &= " dep_loc AS location_id,"
            Sql &= " a.cst_name AS location_name, "
            Sql &= " a.cst_city AS location_city, "
            Sql &= " a.cst_state AS location_state,"
            Sql &= " dep_off AS office, "
            Sql &= " dep_ncal AS next_cal, "
            Sql &= " dep_lcal AS prev_cal,"
            Sql &= " s.name AS status, "
            Sql &= " t.name AS type,"
            Sql &= " name = (dep_manuf + ' ' + dep_mod + ' (' + dep_ser + ')'),"
            Sql &= " date_last_cal = (SELECT TOP 1 so.date_completed FROM work_order wo LEFT OUTER JOIN service_order so ON wo.service_order_id=so.id WHERE calibrated=1 AND equipment_id=equip.dep_id ORDER BY so.date_completed DESC)"
            Sql &= " FROM DEPREC AS equip INNER JOIN"
            Sql &= " ADDRESS a ON equip.dep_loc = a.cst_no LEFT OUTER JOIN"
            Sql &= " equipment_type t ON equip.dep_type = t.id LEFT OUTER JOIN"
            Sql &= " equipment_status s ON dep_stat=s.code INNER JOIN"
            Sql &= " cal_agreement_equipment cae ON equip.dep_id=cae.equipment_id"
            Sql &= " WHERE  cae.agreement_id=" & Me._Id
            Sql &= " ORDER BY name"
            Me.Equipment = Me.Database.GetAll(Sql)
            Me.Equipment.Columns.Add("deleted")
            Me.Equipment.Columns("deleted").DefaultValue = 0
            If Not Me.Database.LastQuery.Successful Then
                Throw New cException(cException.SeverityRating.Serious, "Error getting equipment.", Me.Database.LastQuery.ErrorMsg)
            End If
            ' Send reload event
            RaiseEvent Reload()
        Else
            Dim Err As String = Me.Database.LastQuery.ErrorMsg
        End If
    End Sub

    Public Sub Save()
        If Me._Id = Nothing Then
            Dim Sql As String = "INSERT INTO cal_agreement "
            Sql &= " (ship_to_no, bill_to_no, contact_name, contact_id, purchase_order, frequency_months, frequency_days, price, strict_deadline, comments, date_start,"
            Sql &= " created_by, date_created, date_end, last_updated_by, date_last_updated, appointment_only,"
            Sql &= " estimated_hours, service_area_id, internal_notes, canceled, non_recurring, truck_type_id, formula, date_next)"
            Sql &= " VALUES (@ship_to_no, @bill_to_no, @contact_name, @contact_id, @purchase_order, @frequency_months, @frequency_days, "
            Sql &= Me.Database.ToCurrency("@price") & ", @strict_deadline, @comments,"
            Sql &= " @date_start, @created_by, " & Me.Database.Timestamp & ", @date_end, @created_by, " & Me.Database.Timestamp & ", @appointment_only,"
            Sql &= " @estimated_hours, @service_area_id, @internal_notes, @canceled, @non_recurring, @truck_type_id, @formula, @date_next)"
            Sql = Sql.Replace("@bill_to_no", Me.Database.Escape(Me.BillTo))
            Sql = Sql.Replace("@ship_to_no", Me.Database.Escape(Me.ShipTo))
            Sql = Sql.Replace("@contact_name", Me.Database.Escape(Me.Contact))
            Sql = Sql.Replace("@contact_id", Me.Database.Escape(Me.ContactID))
            Sql = Sql.Replace("@purchase_order", Me.Database.Escape(Me.Po))
            Sql = Sql.Replace("@frequency_months", Me.Database.Escape(Me.FrequencyMonths))
            Sql = Sql.Replace("@frequency_days", Me.Database.Escape(Me.FrequencyDays))
            Sql = Sql.Replace("@price", Me.Database.Escape(Me.Price))
            Sql = Sql.Replace("@strict_deadline", Me.Database.Escape(Me.StrictDeadline))
            Sql = Sql.Replace("@appointment_only", Me.Database.Escape(Me.AppointmentOnly))
            Sql = Sql.Replace("@comments", Me.Database.Escape(Me.ExternalNotes))
            Sql = Sql.Replace("@internal_notes", Me.Database.Escape(Me.InternalNotes))
            Sql = Sql.Replace("@date_start", Me.Database.Escape(Me.DateStart))
            Sql = Sql.Replace("@date_end", Me.Database.Escape(Me.DateToReview))
            Sql = Sql.Replace("@service_area_id", Me.Database.Escape(Me.ServiceAreaId))
            Sql = Sql.Replace("@estimated_hours", Me.Database.Escape(Me.EstimatedHours))
            Sql = Sql.Replace("@canceled", Me.Database.Escape(Me.Canceled))
            Sql = Sql.Replace("@truck_type_id", Me.Database.Escape(Me.TruckTypeId))
            Sql = Sql.Replace("@non_recurring", Me.Database.Escape(Me.NonRecurring))
            Sql = Sql.Replace("@formula", Me.Database.Escape(Me.Formula))
            Sql = Sql.Replace("@created_by", Me.Database.Escape(Me.CreatedBy))
            Sql = Sql.Replace("@date_next", Me.Database.Escape(Me.DateNextCal))
            ' Run insert command
            Me.Database.InsertAndReturnId(Sql)
            If Me.Database.LastQuery.Successful Then
                ' Set id
                Me._Id = Me.Database.LastQuery.InsertId
                ' Add equipment
                For Each dr As DataRow In Me.Equipment.Rows
                    If Not dr.RowState = DataRowState.Deleted Then
                        Dim Sql2 As String = "INSERT INTO cal_agreement_equipment (agreement_id, equipment_id) "
                        Sql2 &= " VALUES (" & Me._Id & ", " & dr.Item("id") & ")"
                        Me.Database.Execute(Sql2)
                    End If
                Next
                Me.SaveLineItems()
                Me.Open(Me._Id)
            Else
                Throw New Exception(Me.Database.LastQuery.ErrorMsg & " " & Sql)
            End If
        Else
            Dim Sql As String = "UPDATE cal_agreement SET "
            Sql &= " contact_name=@contact_name,"
            Sql &= " contact_id=@contact_id,"
            Sql &= " bill_to_no=@bill_to_no,"
            Sql &= " ship_to_no=@ship_to_no,"
            Sql &= " purchase_order=@purchase_order, "
            Sql &= " frequency_months=@frequency_months,"
            Sql &= " frequency_days=@frequency_days,"
            Sql &= " price=" & Me.Database.ToCurrency("@price") & ", "
            Sql &= " strict_deadline=@strict_deadline, "
            Sql &= " comments=@comments, "
            Sql &= " date_start=@date_start, "
            Sql &= " date_end=@date_end, "
            Sql &= " date_next=@date_next,"
            Sql &= " last_updated_by=@updated_by, "
            Sql &= " date_last_updated=" & Me.Database.Timestamp & ","
            Sql &= " appointment_only=@appointment_only,"
            Sql &= " estimated_hours=@estimated_hours,"
            Sql &= " service_area_id=@service_area_id,"
            Sql &= " internal_notes=@internal_notes,"
            Sql &= " canceled=@canceled,"
            Sql &= " formula=@formula,"
            Sql &= " truck_type_id=@truck_type_id,"
            Sql &= " non_recurring=@non_recurring,"
            Sql &= " cancelled_by=@cancelled_by,"
            Sql &= " date_cancel=@date_cancel,"
            Sql &= " cancel_reason=@cancel_reason"
            Sql &= " WHERE id=@agreement_id"
            Sql = Sql.Replace("@agreement_id", Me.Database.Escape(Me._Id))
            Sql = Sql.Replace("@bill_to_no", Me.Database.Escape(Me.BillTo))
            Sql = Sql.Replace("@ship_to_no", Me.Database.Escape(Me.ShipTo))
            Sql = Sql.Replace("@contact_name", Me.Database.Escape(Me.Contact))
            Sql = Sql.Replace("@contact_id", Me.Database.Escape(Me.ContactID))
            Sql = Sql.Replace("@purchase_order", Me.Database.Escape(Me.Po))
            Sql = Sql.Replace("@frequency_months", Me.FrequencyMonths)
            Sql = Sql.Replace("@frequency_days", Me.FrequencyDays)
            Sql = Sql.Replace("@price", Me.Price)
            Sql = Sql.Replace("@strict_deadline", Me.Database.Escape(Me.StrictDeadline))
            Sql = Sql.Replace("@appointment_only", Me.Database.Escape(Me.AppointmentOnly))
            Sql = Sql.Replace("@comments", Me.Database.Escape(Me.ExternalNotes))
            Sql = Sql.Replace("@internal_notes", Me.Database.Escape(Me.InternalNotes))
            Sql = Sql.Replace("@date_start", Me.Database.Escape(Me.DateStart))
            Sql = Sql.Replace("@date_end", Me.Database.Escape(Me.DateToReview))
            Sql = Sql.Replace("@service_area_id", Me.Database.Escape(Me.ServiceAreaId))
            Sql = Sql.Replace("@estimated_hours", Me.Database.Escape(Me.EstimatedHours))
            Sql = Sql.Replace("@canceled", Me.Database.Escape(Me.Canceled))
            Sql = Sql.Replace("@truck_type_id", Me.Database.Escape(Me.TruckTypeId))
            Sql = Sql.Replace("@formula", Me.Database.Escape(Me.Formula))
            Sql = Sql.Replace("@non_recurring", Me.Database.Escape(Me.NonRecurring))
            Sql = Sql.Replace("@updated_by", Me.Database.Escape(Me.LastUpdatedBy))
            Sql = Sql.Replace("@date_cancel", Me.Database.Escape(IIf(Me.DateCancelled = Nothing, DBNull.Value, Me.DateCancelled)))
            Sql = Sql.Replace("@cancel_reason", Me.Database.Escape(Me.CancelReason))
            Sql = Sql.Replace("@cancelled_by", Me.Database.Escape(Me.CancelledBy))
            Sql = Sql.Replace("@date_next", Me.Database.Escape(Me.DateNextCal))
            Me.Database.Execute(Sql)
            If Me.Database.LastQuery.Successful Then
                For Each dr As DataRow In Me.Equipment.Rows
                    If dr.RowState = DataRowState.Added Then
                        If dr.Item("deleted") = 0 Then
                            Dim Sql2 As String = "INSERT INTO cal_agreement_equipment (agreement_id, equipment_id) "
                            Sql2 &= " VALUES (" & Me._Id & ", " & dr.Item("id") & ")"
                            Me.Database.Execute(Sql2)
                        End If
                    ElseIf dr.RowState = DataRowState.Modified Then
                        If dr.Item("deleted") = 1 Then
                            Dim Sql2 As String = "DELETE FROM cal_agreement_equipment "
                            Sql2 &= " WHERE agreement_id=" & Me._Id
                            Sql2 &= " AND equipment_id=" & dr.Item("id")
                            Me.Database.Execute(Sql2)
                        End If
                    End If
                Next
                Me.SaveLineItems()
                RaiseEvent Saved(Me)
                Me.Open(Me._Id)
            Else
                Throw New Exception(Me.Database.LastQuery.ErrorMsg & " " & Sql)
            End If
        End If
    End Sub

    Private Sub SaveLineItems()
        For Each Row As DataRow In Me.LineItems.Rows
            Select Case Row.RowState
                Case DataRowState.Modified
                    If Row.Item("quantity") > 0 Then
                        Dim Sql As String = ""
                        Sql = "UPDATE cal_agreement_item SET"
                        Sql &= " quantity=" & Me.Database.Escape(Row.Item("quantity"))
                        Sql &= ", part_no=" & Me.Database.Escape(Row.Item("part_no"))
                        Sql &= ", description=" & Me.Database.Escape(Row.Item("description"))
                        Sql &= ", price=" & Me.Database.Escape(Row.Item("price"))
                        Sql &= ", tax_status_id=" & Me.Database.Escape(Row.Item("tax_status_id"))
                        Sql &= " WHERE id=" & Row.Item("id")
                        Me.Database.Execute(Sql)
                        If Not Me.Database.LastQuery.Successful Then
                            Dim Msg As String = Me.Database.LastQuery.ErrorMsg
                            MessageBox.Show(Msg & " - " & Sql)
                        End If
                    Else
                        If Row.Item("id") IsNot DBNull.Value Then
                            Me.Database.Execute("DELETE FROM cal_agreement_item WHERE id=" & Row.Item("id"))
                        End If
                    End If
                Case DataRowState.Added
                    If Row.Item("quantity") > 0 Then
                        Dim Sql As String = ""
                        Sql = "INSERT INTO cal_agreement_item"
                        Sql &= " (cal_agreement_id, quantity, part_no, description, price, tax_status_id)"
                        Sql &= " VALUES (" & Me.Id
                        Sql &= ", " & Me.Database.Escape(Row.Item("quantity"))
                        Sql &= ", " & Me.Database.Escape(Row.Item("part_no"))
                        Sql &= ", " & Me.Database.Escape(Row.Item("description"))
                        Sql &= ", " & Me.Database.ToCurrency(Me.Database.Escape(Row.Item("price")))
                        Sql &= ", " & Me.Database.Escape(Row.Item("tax_status_id"))
                        Sql &= ")"
                        Me.Database.InsertAndReturnId(Sql)
                        If Not Me.Database.LastQuery.Successful Then
                            Dim Msg As String = Me.Database.LastQuery.ErrorMsg
                            MessageBox.Show(Msg & " - " & Sql)
                        End If
                    End If
            End Select
        Next
    End Sub

    Private Function IsNull(ByVal Value As Object, ByVal DefaultVal As Object) As Object
        If Value IsNot DBNull.Value Then
            Return Value
        Else
            Return DefaultVal
        End If
    End Function

    Private Sub CreateBlankEquipmentTable()
        Me.Equipment = New DataTable
        Me.Equipment.Columns.Add("id")
        Me.Equipment.Columns.Add("name")
        Me.Equipment.Columns.Add("date_last_cal")
        Me.Equipment.Columns.Add("deleted")
        Me.Equipment.Columns("deleted").DefaultValue = 0
    End Sub

    Public Function ServiceOrders() As DataTable
        Dim Sql As String = "SELECT so.id AS service_order_id, so.contact_id, so.location_id, so.office, so.cal_agreement_id, "
        Sql &= " so.notes, so.date_due, so.date_scheduled, so.date_completed, so.invoice_id, so.truck_type_id, "
        Sql &= " so.rate_type_id, so.work_type_id, so.contract, so.techs, so.helpers, "
        Sql &= " so.approved_name, so.return_trip_required,  so.date_created, so.created_by,"
        Sql &= " contact_name, approved_name, caller_name,"
        Sql &= " CASE "
        Sql &= " WHEN so.voided=1 THEN 'Voided'"
        Sql &= " WHEN so.invoice_id > 0 THEN 'Invoiced'"
        Sql &= " WHEN so.date_completed IS NOT NULL THEN 'Completed'"
        Sql &= " WHEN so.date_scheduled IS NOT NULL THEN 'Scheduled'"
        Sql &= " ELSE 'Received'"
        Sql &= " END AS stage"
        Sql &= " FROM service_order so"
        Sql &= " WHERE so.cal_agreement_id=" & Me._Id
        Sql &= " ORDER BY so.date_created DESC"
        Return Me.Database.GetAll(Sql)
    End Function

    Public Function FreeEquipment(Optional ByVal CustomerNo As String = "", Optional ByVal NotOnAnotherActiveAgreement As Boolean = True) As DataTable
        If CustomerNo.Length = 0 Then
            CustomerNo = Me.ShipTo
        End If
        Dim Sql As String = "SELECT equip.*, eqtype.name AS type,"
        Sql &= " name = CASE WHEN equip.dep_ser IS NOT NULL THEN equip.dep_manuf + ' ' + equip.dep_mod  + ' (' + equip.dep_ser + ')'"
        Sql &= " ELSE equip.dep_manuf + ' ' + equip.dep_mod END"
        Sql &= " FROM DEPREC AS equip INNER JOIN"
        Sql &= " ADDRESS a ON equip.dep_loc = a.cst_no LEFT OUTER JOIN"
        Sql &= " equipment_type eqtype ON equip.dep_type = eqtype.id"
        Sql &= " WHERE a.cst_no=" & Me.Database.Escape(CustomerNo)
        Sql &= " AND equip.inactive=0"
        If NotOnAnotherActiveAgreement Then
            Sql &= " AND equip.dep_id NOT IN (SELECT cae.equipment_id "
            Sql &= " FROM cal_agreement_equipment cae"
            Sql &= " JOIN cal_agreement AS ca ON ca.id=cae.agreement_id"
            Sql &= " WHERE ca.canceled=0)"
        End If
        Sql &= " ORDER BY name"
        Return Me.Database.GetAll(Sql)
    End Function

    Public Function TruckTypes() As DataTable
        Return Me.Database.GetAll("SELECT id, name FROM truck_type ORDER BY name")
    End Function

    Public Function ServiceAreas() As DataTable
        Return Me.Database.GetAll("SELECT id, name FROM service_area ORDER BY name")
    End Function

    Public Function GetNextDueDate(Optional ByVal Previous As DateTime = Nothing) As DateTime
        If Previous = Nothing Then
            ' Get last date
            If Me.Settings.GetValue("Cal Agreement Due Key", "date_due") = "date_due" Then
                Previous = Me.DatePreviousCalDue
            Else
                Previous = Me.DatePreviousCal
            End If
        End If
        ' If there was a last date
        If Previous <> Nothing Then
            ' Calculate next date
            Dim nd As DateTime
            If Me.FrequencyMonths > 0 Then
                nd = Previous.AddMonths(Me.FrequencyMonths)
            Else
                nd = Previous.AddDays(Me.FrequencyDays)
            End If
            ' Determine if next date needs to be adjusted or not and return it
            If Me.StrictDeadline Then
                Return nd
            Else
                If nd.Month = 12 Then
                    nd = New Date(nd.Year, nd.Month, 31)
                Else
                    nd = New Date(nd.Year, nd.Month + 1, 1)
                    nd = nd.AddDays(-1)
                End If
                Return nd
            End If
        Else
            Return Me.DateStart
        End If
    End Function

End Class
