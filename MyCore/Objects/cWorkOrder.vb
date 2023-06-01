Public Class cWorkOrder

    Dim _Id As Integer = 0
    Dim _ServiceOrderId As Integer = 0
    Dim _Equipment As cEquipment = Nothing
    Dim Database As MyCore.Data.EasySql

    Dim _EquipmentId As Integer = 0
    Public Description As String = ""
    Public ProblemReported As String = ""
    Public ProblemFound As String = ""
    Public CorrectiveAction As String = ""
    Public Calibrated As Boolean = False
    Public Skipped As Boolean = False
    Public ReportToState As Boolean = False
    Public NeedsReturnService As Boolean = False
    Public Temperature As String = ""
    Public Humidity As String = ""
    Public Barometer As String = ""
    Public CalibrationTimestamp As DateTime = Nothing
    Public DateNextDue As Date = Nothing
    Public EnvironmentalCondition As Integer = 0
    Public CustomerStandards As String = ""
    Public CalibrationProcedure As Integer = 1
    Public EquipmentCondition As Integer = 0
    Public FoundInTolerance As Boolean = False
    Public UncertaintyStatementId As Integer = 0
    Public UncertaintyValue As Double = 0
    Public UncertaintyUnits As Integer = 0
    Public Technician As String = ""
    Public CalibrationTests As DataTable

    Public LastUpdatedBy As String = ""
    Public CreatedBy As String = ""

    Public Standards As DataTable
    Public WorkOrderItems As DataTable

    Public Event Reload()

    Public Property EquipmentId() As Integer
        Get
            Return Me._EquipmentId
        End Get
        Set(ByVal value As Integer)
            Me._EquipmentId = value
            Me.LoadEquipment()
        End Set
    End Property

    Public ReadOnly Property Id() As Integer
        Get
            Return Me._Id
        End Get
    End Property

    Public ReadOnly Property ServiceOrderId() As Integer
        Get
            Return Me._ServiceOrderId
        End Get
    End Property

    Public ReadOnly Property Equipment() As cEquipment
        Get
            Return Me._Equipment
        End Get
    End Property


    Public ReadOnly Property TravelTotal() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.WorkOrderItems.Rows
                If Row.Item("item_type_id") IsNot DBNull.Value Then
                    If Row.Item("item_type_id") = 3 Then
                        Amt += Row.Item("quantity") * Row.Item("unit_price")
                    End If
                End If
            Next
            Return Amt
        End Get
    End Property

    Public ReadOnly Property LaborTotal() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.WorkOrderItems.Rows
                If Row.Item("item_type_id") IsNot DBNull.Value Then
                    If Row.Item("item_type_id") = 2 Then
                        Amt += Row.Item("quantity") * Row.Item("unit_price")
                    End If
                End If
            Next
            Return Amt
        End Get
    End Property

    Public ReadOnly Property PartsTotal() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.WorkOrderItems.Rows
                If Row.Item("item_type_id") IsNot DBNull.Value Then
                    If Row.Item("item_type_id") = 1 Then
                        Amt += Row.Item("quantity") * Row.Item("unit_price")
                    End If
                End If
            Next
            Return Amt
        End Get
    End Property

    Public ReadOnly Property LineItemsTotal() As Double
        Get
            Dim Amt As Double = 0
            For Each Row As DataRow In Me.WorkOrderItems.Rows
                If Row.RowState <> DataRowState.Deleted And Row.RowState <> DataRowState.Detached Then
                    Dim qty As Double = IIf(Row.Item("quantity") Is DBNull.Value, 1, Row.Item("quantity"))
                    Dim price As Double = IIf(Row.Item("unit_price") Is DBNull.Value, 0, Row.Item("unit_price"))
                    Amt += qty * price
                End If
            Next
            Return Amt
        End Get
    End Property

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
    End Sub

    Public Function Open(ByVal Id As Integer) As Boolean
        Dim Sql As String = "SELECT wo.*,"
        Sql &= " equip.dep_manuf, equip.dep_cap, equip.dep_ser, equip.dep_mod, equip.dep_countby"
        Sql &= " FROM work_order wo"
        Sql &= " LEFT OUTER JOIN DEPREC equip ON equip.dep_id=wo.equipment_id"
        Sql &= " WHERE wo.id=" & Id
        Dim Row As DataRow = Me.Database.GetRow(Sql)
        If Me.Database.LastQuery.RowsReturned = 1 Then
            Me.EquipmentId = IIf(Row.Item("equipment_id") Is DBNull.Value, 0, Row.Item("equipment_id"))
            Me.Description = Row.Item("description")
            Me.ProblemReported = Row.Item("problem_reported")
            Me.ProblemFound = Row.Item("problem_found")
            Me.CorrectiveAction = Row.Item("corrective_action")
            Me.Calibrated = Row.Item("calibrated")
            Me.ReportToState = Row.Item("notify_state")
            Me.Skipped = Row.Item("skipped")
            Me._ServiceOrderId = Row.Item("service_order_id")
            Me.NeedsReturnService = Row.Item("needs_return_service")
            Me.Temperature = Me.IsNull(Row.Item("temperature"), "")
            Me.Humidity = Me.IsNull(Row.Item("humidity"), "")
            Me.Barometer = Me.IsNull(Row.Item("barometer"), "")
            Me.CalibrationTimestamp = Me.IsNull(Row.Item("calibration_timestamp"), Nothing)
            Me.DateNextDue = Me.IsNull(Row.Item("date_next_due"), Nothing)
            Me.EnvironmentalCondition = Row.Item("environmental_condition")
            Me.CustomerStandards = Row.Item("customer_standards")
            Me.CalibrationProcedure = Row.Item("calibration_procedure")
            Me.EquipmentCondition = Row.Item("equipment_condition")
            Me.FoundInTolerance = IIf(Row.Item("found_in_tolerance") Is DBNull.Value, False, Row.Item("found_in_tolerance"))
            Me.Technician = Row.Item("technician")
            Me.UncertaintyStatementId = Row.Item("uncertainty_statement_id")
            Me.UncertaintyValue = Row.Item("uncertainty_value")
            Me.UncertaintyUnits = Row.Item("uncertainty_units")
            Me._Id = Id
            Me.CalibrationTests = Me.GetCalibrationTests
            Me.LoadEquipment()
            Me.PopulateWorkOrderItems()
            Me.PopulateStandardsUsed()
            RaiseEvent Reload()
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub LoadEquipment()
        If Me._EquipmentId > 0 Then
            Me._Equipment = New cEquipment(Me.Database)
            Me._Equipment.Open(Me._EquipmentId)
        Else
            Me._Equipment = Nothing
        End If
    End Sub

    Private Sub PopulateStandardsUsed()
        Dim Sql As String = "SELECT asset_no, id FROM standards_used WHERE work_order_id=" & Me.Database.Escape(Me._Id)
        Me.Standards = Me.Database.GetAll(Sql)
        Me.Standards.Columns.Add("delete")
    End Sub

    Public Sub PopulateWorkOrderItems()
        Dim Sql As String = "SELECT i.*, i.unit_price AS unit, (i.quantity * i.unit_price) AS ext_price,"
        Sql &= " im.item_type_id"
        Sql &= " FROM work_order_item i LEFT OUTER JOIN item_master im ON i.part_no=im.part_no"
        Sql &= " WHERE work_order_id=" & Me._Id
        Me.WorkOrderItems = Me.Database.GetAll(Sql)
    End Sub

    Public Sub Save()
        If Me._Id > 0 Then
            Dim Sql As String = "UPDATE work_order"
            Sql &= " SET service_order_id=@service_order_id, equipment_id=@equipment_id,"
            Sql &= " description=@description, problem_reported=@problem_reported,"
            Sql &= " problem_found=@problem_found, calibrated=@calibrated, notify_state=@notify_state,"
            Sql &= " corrective_action=@corrective_action, skipped=@skipped,"
            Sql &= " date_last_updated=" & Me.Database.Timestamp & ","
            Sql &= " equipment_condition=@equipment_condition,"
            Sql &= " environmental_condition=@environmental_condition,"
            Sql &= " needs_return_service=@needs_return_service,"
            Sql &= " temperature=@temperature,"
            Sql &= " humidity=@humidity,"
            Sql &= " barometer=@barometer,"
            Sql &= " calibration_timestamp=@calibration_timestamp,"
            Sql &= " date_next_due=@date_next_due,"
            Sql &= " customer_standards=@customer_standards,"
            Sql &= " calibration_procedure=@calibration_procedure,"
            Sql &= " found_in_tolerance=@found_in_tolerance"
            Sql &= " WHERE id=@work_order_id"
            Sql = Sql.Replace("@work_order_id", Me._Id)
            Sql = Sql.Replace("@description", Me.Database.Escape(Me.Description))
            Sql = Sql.Replace("@problem_reported", Me.Database.Escape(Me.ProblemReported))
            Sql = Sql.Replace("@problem_found", Me.Database.Escape(Me.ProblemFound))
            Sql = Sql.Replace("@corrective_action", Me.Database.Escape(Me.CorrectiveAction))
            Sql = Sql.Replace("@equipment_id", Me.EquipmentId)
            Sql = Sql.Replace("@calibrated", Me.Database.Escape(Me.Calibrated))
            Sql = Sql.Replace("@notify_state", Me.Database.Escape(Me.ReportToState))
            Sql = Sql.Replace("@skipped", Me.Database.Escape(Me.Skipped))
            Sql = Sql.Replace("@service_order_id", Me._ServiceOrderId)
            Sql = Sql.Replace("@equipment_condition", Me.EquipmentCondition)
            Sql = Sql.Replace("@environmental_condition", Me.EnvironmentalCondition)
            Sql = Sql.Replace("@needs_return_service", Me.Database.Escape(Me.NeedsReturnService))
            Sql = Sql.Replace("@temperature", Me.Database.Escape(Me.Temperature))
            Sql = Sql.Replace("@humidity", Me.Database.Escape(Me.Humidity))
            Sql = Sql.Replace("@barometer", Me.Database.Escape(Me.Barometer))
            Sql = Sql.Replace("@calibration_timestamp", Me.Database.Escape(IIf(Me.CalibrationTimestamp = Nothing, DBNull.Value, Me.CalibrationTimestamp)))
            Sql = Sql.Replace("@date_next_due", Me.Database.Escape(IIf(Me.DateNextDue = Nothing, DBNull.Value, Me.DateNextDue)))
            Sql = Sql.Replace("@customer_standards", Me.Database.Escape(Me.CustomerStandards))
            Sql = Sql.Replace("@calibration_procedure", Me.CalibrationProcedure)
            Sql = Sql.Replace("@found_in_tolerance", Me.Database.Escape(Me.FoundInTolerance))
            Me.Database.Execute(Sql)
            If Me.Database.LastQuery.Successful Then
                ' Calibration tests
                For Each Row As DataRow In Me.CalibrationTests.Rows
                    Sql = ""
                    If Row.RowState = DataRowState.Added Then
                        Sql = "INSERT INTO calibration_test (work_order_id, instrument, test_name, applied,"
                        Sql &= " initial, final, pass, as_found_pass, tolerance_value, note, date_last_updated"
                        Sql &= " )"
                        Sql &= " VALUES (@work_order_id, @instrument, @test_name, @applied, @initial, @final, @pass, @as_found_pass, @tolerance_value, @note, " & Me.Database.Timestamp & ")"
                    ElseIf Row.RowState = DataRowState.Modified Then
                        Sql = "UPDATE calibration_test SET "
                        Sql &= " work_order_id=@work_order_id, instrument=@instrument,"
                        Sql &= " test_name=@test_name, applied=@applied, initial=@initial,"
                        Sql &= " final=@final, pass=@pass, note=@note, tolerance_value=@tolerance_value,"
                        Sql &= " as_found_pass=@as_found_pass, date_last_updated = " & Me.Database.Timestamp & ""
                        Sql &= " WHERE id=@id"
                    End If
                    If Sql.Length > 0 Then
                        If Row.Item("id") IsNot DBNull.Value Then
                            Sql = Sql.Replace("@id", Row.Item("id"))
                        End If
                        Sql = Sql.Replace("@work_order_id", Me._Id)
                        Sql = Sql.Replace("@instrument", Me.Database.Escape(Row.Item("instrument")))
                        Sql = Sql.Replace("@test_name", Me.Database.Escape(Row.Item("test_name")))
                        Sql = Sql.Replace("@applied", Me.Database.Escape(Row.Item("applied")))
                        Sql = Sql.Replace("@initial", Me.Database.Escape(Row.Item("initial")))
                        Sql = Sql.Replace("@final", Me.Database.Escape(Row.Item("final")))
                        Sql = Sql.Replace("@pass", Me.Database.Escape(IIf(Row.Item("pass") Is DBNull.Value, False, Row.Item("pass"))))
                        Sql = Sql.Replace("@as_found_pass", Me.Database.Escape(IIf(Row.Item("as_found_pass") Is DBNull.Value, False, Row.Item("as_found_pass"))))
                        Sql = Sql.Replace("@tolerance_value", Me.Database.Escape(Row.Item("tolerance_value")))
                        Sql = Sql.Replace("@note", Me.Database.Escape(Row.Item("note")))
                        Me.Database.Execute(Sql)
                    End If
                Next
                ' Additional charges
                Dim Added As Integer = 0
                Dim Modified As Integer = 0
                Dim Other As Integer = 0
                For Each Row As DataRow In Me.WorkOrderItems.Rows
                    Sql = ""
                    If Row.RowState = DataRowState.Added Then
                        If Row.Item("quantity") Is DBNull.Value Then
                            Row.Item("quantity") = 1
                        End If
                        If Row.Item("quantity") > 0 Then
                            Sql = "INSERT INTO work_order_item (work_order_id, quantity, unit_price, description,"
                            Sql &= " date_last_updated, part_no, serial_no, tax_status_id, station_id, equipment_id)"
                            Sql &= " VALUES (@work_order_id, @quantity, " & Me.Database.ToCurrency("@unit_price")
                            Sql &= ", @description, " & Me.Database.Timestamp & ", @part_no, @serial_no, @tax_status_id,"
                            Sql &= " @station_id, @equipment_id)"
                        End If
                        Added += 1
                    ElseIf Row.RowState = DataRowState.Modified Then
                        If Row.Item("quantity") Is DBNull.Value Then
                            Row.Item("quantity") = 0
                        End If
                        If Row.Item("quantity") > 0 Then
                            Sql = "UPDATE work_order_item "
                            Sql &= " SET quantity=@quantity, unit_price=" & Me.Database.ToCurrency("@unit_price") & ", description=@description, "
                            Sql &= " work_order_id=@work_order_id, date_last_updated=" & Me.Database.Timestamp & ", part_no=@part_no, serial_no=@serial_no,"
                            Sql &= " tax_status_id=@tax_status_id, station_id=@station_id, equipment_id=@equipment_id"
                            Sql &= " WHERE id=@id"
                        Else
                            Sql = "DELETE FROM work_order_item WHERE id=@id"
                        End If
                        Modified += 1
                    Else
                        Dim State As String = Row.RowState
                        Other += 1
                    End If
                    If Sql.Length > 0 Then
                        If Row.Item("id") IsNot DBNull.Value Then
                            Sql = Sql.Replace("@id", Row.Item("id"))
                        End If
                        Sql = Sql.Replace("@quantity", Me.Database.Escape(Row.Item("quantity")))
                        Sql = Sql.Replace("@description", Me.Database.Escape(Me.IsNull(Row.Item("description"))))
                        Sql = Sql.Replace("@unit_price", Me.Database.Escape(Row.Item("unit_price")))
                        Sql = Sql.Replace("@part_no", Me.Database.Escape(Me.IsNull(Row.Item("part_no"))))
                        Sql = Sql.Replace("@serial_no", Me.Database.Escape(Me.IsNull(Row.Item("serial_no"))))
                        Sql = Sql.Replace("@work_order_id", Me._Id)
                        Sql = Sql.Replace("@tax_status_id", Me.Database.Escape(Me.IsNull(Row.Item("tax_status_id"))))
                        Sql = Sql.Replace("@station_id", Me.Database.Escape(Me.IsNull(Row.Item("station_id"), 0)))
                        Sql = Sql.Replace("@equipment_id", Me.Database.Escape(Me.IsNull(Row.Item("equipment_id"), 0)))
                        Me.Database.Execute(Sql)
                        If Not Me.Database.LastQuery.Successful Then
                            MsgBox(Me.Database.LastQuery.ErrorMsg)
                        End If
                    End If
                Next
                ' Standards used
                For Each Row As DataRow In Me.Standards.Rows
                    If Row.RowState = DataRowState.Added Then
                        Sql = "INSERT INTO standards_used (work_order_id, asset_no, date_created)"
                        Sql &= " VALUES (@work_order_id, @asset_no, @date_created)"
                        Sql = Sql.Replace("@work_order_id", Me.Database.Escape(Me._Id))
                        Sql = Sql.Replace("@asset_no", Me.Database.Escape(Row.Item("asset_no")))
                        Sql = Sql.Replace("@date_created", Me.Database.Escape(Now))
                        Me.Database.Execute(Sql)
                        If Not Me.Database.LastQuery.Successful Then
                            MsgBox(Me.Database.LastQuery.ErrorMsg)
                        End If
                    ElseIf Row.RowState = DataRowState.Modified Then
                        If Row.Item("delete") IsNot DBNull.Value And Row.Item("id") IsNot DBNull.Value Then
                            If Row.Item("delete") = 1 Then
                                Sql = "DELETE FROM standards_used WHERE id=" & Row.Item("id")
                                Me.Database.Execute(Sql)
                            End If
                        End If
                    End If
                Next

                Me.Open(Me._Id)

            Else
                Throw New Exception(Me.Database.LastQuery.ErrorMsg)
            End If
        Else
            ' Save as new work order
            ' no ability for this yet
        End If

    End Sub

    Private Function IsNull(ByVal Value As Object, Optional ByVal DefaultVal As Object = "") As Object
        If Value Is DBNull.Value Then
            Return DefaultVal
        Else
            Return Value
        End If
    End Function

    Public Function Notes() As DataTable
        Dim Sql As String = "SELECT [id], [author], [date_created], [note] FROM work_order_note"
        Sql &= " WHERE work_order_id=" & Me.Id & " ORDER BY date_created ASC"
        Return Me.Database.GetAll(Sql)
    End Function

    Public Sub AddNote(ByVal Note As String)
        Dim Sql As String = "INSERT INTO work_order_note (author, date_created, work_order_id, [note])"
        Sql &= " VALUES (" & Me.Database.Escape(Me.LastUpdatedBy) & ", "
        Sql &= Me.Database.Escape(Now) & ", "
        Sql &= Me.Id & ", "
        Sql &= Me.Database.Escape(Note) & ")"
        Me.Database.Execute(Sql)
    End Sub

    Public Function GetCertTemplate(ByVal DefaultTemplate As String) As String
        ' From equipment
        Dim Id As Integer = 0
        If Me.EquipmentId > 0 Then
            Id = Me.Database.GetOne("SELECT cert_template_id FROM DEPREC WHERE dep_id=" & Me.EquipmentId)
            If Id > 0 Then
                Return Me.Database.GetOne("SELECT html FROM template WHERE id=" & Id)
            End If
        End If
        ' From Company
        Dim Sql As String = "SELECT cert_template_id FROM ADDRESS WHERE cst_no=(SELECT location_id FROM service_order WHERE id=" & Me.ServiceOrderId & ")"
        Id = Me.Database.GetOne(Sql)
        If Me.Database.LastQuery.Successful Then
            If Id > 0 Then
                Return Me.Database.GetOne("SELECT html FROM template WHERE id=" & Id)
            End If
        Else
            Dim Msg As String = Me.Database.LastQuery.ErrorMsg
        End If
        ' Default
        Return DefaultTemplate
    End Function

    Public Function ToGravityDocument(ByVal Template As String, Optional ByVal Blank As Boolean = False) As GravityDocument.gDocument
        ' If no template specified
        If Template.Length = 0 Then
            Dim id As Integer = Me.Database.GetOne("SELECT value FROM settings WHERE property='Template Cert'")
            Template = Me.Database.GetOne("SELECT html FROM template WHERE id=" & id)
        End If
        ' Create Gravity Document
        Dim Doc As New GravityDocument.gDocument(Me.Database.GetOne("SELECT value FROM settings WHERE property='Page Height in Pixels'"))
        Doc.LoadXml(Template)
        ' Settings
        Doc.FormType = GravityDocument.gDocument.FormTypes.CalCert
        Doc.ReferenceID = Me.Id
        Doc.LoadXml(Template)
        Dim Page As GravityDocument.gPage = Doc.GetPage(1)

        Dim dt As DataTable
        Dim r As DataRow
        Dim Params As Collection

        Dim so As New cServiceOrder(Me.Database)
        so.Open(Me.ServiceOrderId)

        Page.AddVariable("%service_order_id%", Me.ServiceOrderId)
        Page.AddVariable("%work_order_id%", Me.Id)

        ' Ship To
        Page.AddVariable("%customer_no%", so.ShipToNo)
        Page.AddVariable("%company_name%", so.ShipTo.Name)
        If Not so.ShipTo.Address2.Length = 0 Then
            Page.AddVariable("%address%", so.ShipTo.Address1 & ControlChars.CrLf & so.ShipTo.Address2)
        Else
            Page.AddVariable("%address%", so.ShipTo.Address1)
        End If
        Page.AddVariable("%city%", so.ShipTo.City)
        Page.AddVariable("%state%", so.ShipTo.State)
        Page.AddVariable("%zip%", so.ShipTo.Zip)
        Page.AddVariable("%phone%", so.ShipTo.Phone)
        Page.AddVariable("%fax%", so.ShipTo.Fax)
        Page.AddVariable("%ship_to_no%", so.ShipToNo)

        ' Bill To
        Page.AddVariable("%bill_to_no%", so.BillToNo)
        Page.AddVariable("%bill_to_name%", so.BillTo.Name)
        If so.BillTo.Address2.Length = 0 Then
            Page.AddVariable("%bill_to_address%", so.BillTo.Address1)
        Else
            Page.AddVariable("%bill_to_address%", so.BillTo.Address1 & ControlChars.CrLf & so.BillTo.Address2)
        End If
        Page.AddVariable("%bill_to_city%", so.BillTo.City)
        Page.AddVariable("%bill_to_state%", so.BillTo.State)
        Page.AddVariable("%bill_to_zip%", so.BillTo.Zip)
        Page.AddVariable("%bill_to_phone%", so.BillTo.Phone)
        Page.AddVariable("%bill_to_fax%", so.BillTo.Fax)
        Page.AddVariable("%terms%", so.BillTo.Terms)

        ' Service Order Details
        Page.AddVariable("%service_order_id%", Me.ServiceOrderId)
        Page.AddVariable("%caller%", so.CalledInBy)
        Page.AddVariable("%contact_name%", so.ContactName)
        Page.AddVariable("%contact%", so.ContactName)
        Page.AddVariable("%approved_by%", so.ApprovedBy)
        Page.AddVariable("%date_received%", Format(so.DateCreated, "MM/dd/yyyy"))
        'Page.AddVariable("%received_by%", SO.Item("taken_by_)
        Page.AddVariable("%date_due%", Format(so.DateDue, "MM/dd/yyyy"))
        Page.AddVariable("%po%", so.Po)
        If so.DateScheduled = Nothing Then
            Page.AddVariable("%date_scheduled%", "--")
        Else
            Page.AddVariable("%date_scheduled%", Format(so.DateScheduled, "MM/dd/yyyy"))
        End If
        If so.DateCompleted = Nothing Then
            Page.AddVariable("%date_completed%", "")
            Page.AddVariable("%date%", Now.ToString("MM/dd/yyyy"))
        Else
            Page.AddVariable("%date_completed%", Format(so.DateCompleted, "MM/dd/yyyy"))
            Page.AddVariable("%date%", Format(so.DateCompleted, "MM/dd/yyyy"))
        End If
        Page.AddVariable("%notes%", so.Notes)
        Page.AddVariable("%work_location%", IIf(so.Shop, "In Shop", "On Site"))
        Page.AddVariable("%shop_work_yn%", IIf(so.Shop, "Y", "N"))
        Page.AddVariable("%shop_work_x%", IIf(so.Shop, "X", ""))
        Page.AddVariable("%flat_fee%", "")
        Page.AddVariable("%trip_fee%", "")
        Page.AddVariable("%zone%", "")
        Page.AddVariable("%hourly_rate%", "")

        ' Get attached equipment
        Dim strAttached As String = ""
        Dim Sql As String = ""
        If Me._Equipment IsNot Nothing Then
            Sql = "SELECT DISTINCT equip.dep_id AS id,"
            Sql &= " equip.dep_loc AS location_id,"
            Sql &= " a.cst_name AS location_name, "
            Sql &= " a.cst_city AS location_city, "
            Sql &= " a.cst_state AS location_state,"
            Sql &= " equip.dep_off AS office, "
            Sql &= " equip.dep_ncal AS next_cal, "
            Sql &= " equip.dep_lcal AS prev_cal,"
            Sql &= " s.name AS status, "
            Sql &= " t.name AS type,"
            Sql &= " name = (equip.dep_manuf + ' ' + equip.dep_mod + ' (' + equip.dep_ser + ')')"
            Sql &= " FROM DEPREC AS equip INNER JOIN"
            Sql &= " ADDRESS a ON equip.dep_loc = a.cst_no LEFT OUTER JOIN"
            Sql &= " equipment_type t ON equip.dep_type = t.id LEFT OUTER JOIN"
            Sql &= " equipment_status s ON dep_stat=s.code"
            Sql &= " WHERE"
            Sql &= " (equip.dep_indicator = @indicator_id OR equip.dep_id = @indicator_id)"
            Sql &= " AND dep_id != @self_id"
            Sql &= " ORDER BY name"
            Sql = Sql.Replace("@self_id", Me.EquipmentId)
            If Me._Equipment.AttachedTo > 0 Then
                Sql = Sql.Replace("@indicator_id", Me._Equipment.AttachedTo)
            Else
                Sql = Sql.Replace("@indicator_id", Me.EquipmentId)
            End If
            dt = Me.Database.GetAll(Sql)
            For Each r In dt.Rows
                strAttached &= r.Item("name") & ", "
            Next
            If strAttached.Length > 2 Then
                strAttached = strAttached.Substring(0, strAttached.Length - 2)
            End If
        End If


        ' Equipment details
        If Me._Equipment IsNot Nothing Then
            Page.AddVariable("%equipment_id%", Me.EquipmentId)
            Page.AddVariable("%manufacturer%", Me._Equipment.Manufacturer)
            Page.AddVariable("%model%", Me.Equipment.Model)
            Page.AddVariable("%serial%", Me.Equipment.Serial)
            Page.AddVariable("%asset_no%", Me.Equipment.AssetNo)
            Page.AddVariable("%capacity%", Me.Equipment.Capacity)
            Page.AddVariable("%count_by%", Me.Equipment.CountBy)
            Page.AddVariable("%grad%", Me.Equipment.CountBy)
            Page.AddVariable("%customer_desc%", Me.Equipment.CustomerDescription)
            Page.AddVariable("%count_by_units%", Me.Equipment.CountByUnitsName)
            Page.AddVariable("%grad_units%", Me.Equipment.CountByUnitsName)
            Page.AddVariable("%capacity_units%", Me.Equipment.CapacityUnitsName)
            Page.AddVariable("%tolerance%", Me.Equipment.Tolerenace)
        Else
            Page.AddVariable("%equipment_id%", "")
            Page.AddVariable("%manufacturer%", "")
            Page.AddVariable("%model%", "")
            Page.AddVariable("%serial%", "")
            Page.AddVariable("%asset_no%", "")
            Page.AddVariable("%capacity%", "")
            Page.AddVariable("%capacity_units%", "")
            Page.AddVariable("%count_by%", "")
            Page.AddVariable("%count_by_units%", "")
            Page.AddVariable("%grad%", "")
            Page.AddVariable("%grad_units%", "")
            Page.AddVariable("%tolerance%", "")
            Page.AddVariable("%customer_desc%", "")
        End If
        Page.AddVariable("%problem_found%", Me.ProblemFound)
        Page.AddVariable("%problem_reported%", Me.ProblemReported)
        Page.AddVariable("%corrective_action%", Me.CorrectiveAction)
        Page.AddVariable("%attached_equipment%", strAttached)
        Page.AddVariable("%description%", Me.Description)

        ' Calibrated
        If Not Blank Then
            If Me.Calibrated Then
                Page.AddVariable("%calibrated_yesno%", "Yes")
                Page.AddVariable("%calibrated_yn%", "Y")
                Page.AddVariable("%calibrated_x%", "X")
            Else
                Page.AddVariable("%calibrated_yesno%", "No")
                Page.AddVariable("%calibrated_yn%", "N")
                Page.AddVariable("%calibrated_x%", "")
            End If
            If Me.FoundInTolerance Then
                Page.AddVariable("%fit_yesno%", "Yes")
                Page.AddVariable("%fit_yn%", "Y")
                Page.AddVariable("%fit_x%", "X")
            Else
                Page.AddVariable("%fit_yesno%", "No")
                Page.AddVariable("%fit_yn%", "N")
                Page.AddVariable("%fit_x%", "")
            End If
        Else
            Page.AddVariable("%calibrated_yesno%", "")
            Page.AddVariable("%calibrated_yn%", "")
            Page.AddVariable("%calibrated_x%", "")
            Page.AddVariable("%fit_yesno%", "")
            Page.AddVariable("%fit_yn%", "")
            Page.AddVariable("%fit_x%", "")
            Page.AddVariable("%lit_yesno%", "")
            Page.AddVariable("%lit_yn%", "")
            Page.AddVariable("%lit_x%", "")
        End If

        ' Standards used

        Dim StandardsCount As Integer = 1
        Dim StandardsUsed As DataTable = Me.GetStandardsUsed
        If StandardsUsed.Rows.Count > 0 Then
            Dim strStandards As String = ""
            For i As Integer = 0 To StandardsUsed.Rows.Count - 1
                If StandardsUsed.Rows(i).Item("asset_no").ToString.Length > 0 Then
                    strStandards &= StandardsUsed.Rows(i).Item("asset_no") & ", "
                End If
            Next
            If strStandards.Length > 2 Then
                strStandards = strStandards.Substring(0, strStandards.Length - 2)
            End If
            Page.AddVariable("%standards_used%", strStandards)
            ' Build new test data table
            Dim NewStandardsTable As New DataTable
            NewStandardsTable.Columns.Add("asset_no")
            NewStandardsTable.Columns.Add("serial_no")
            NewStandardsTable.Columns.Add("test_no")
            NewStandardsTable.Columns.Add("date_tested")
            NewStandardsTable.Columns.Add("date_expires")
            NewStandardsTable.Columns.Add("uncertainty")
            ' Populate values
            For Each Row As DataRow In StandardsUsed.Rows
                If Me.IsNull(Row.Item("asset_no")).ToString.Length > 0 Then
                    Dim InternalDateTested As DateTime = Me.IsNull(Row.Item("date_last_internal_cal"), New DateTime(1970, 1, 1))
                    Dim InternalDateExpires As DateTime = InternalDateTested.AddMonths(Row.Item("internal_cal_frequency"))
                    Dim ExternalDateTested As DateTime = Me.IsNull(Row.Item("date_tested"), New DateTime(1970, 1, 1))
                    Dim ExternalDateExpires As DateTime = Me.IsNull(Row.Item("date_expires"), New DateTime(1970, 1, 1))
                    Dim nr As DataRow = NewStandardsTable.NewRow
                    nr.Item("asset_no") = Row.Item("asset_no")
                    nr.Item("serial_no") = Row.Item("serial_no")
                    Page.AddVariable("%su_ano" & StandardsCount & "%", Me.IsNull(Row.Item("asset_no")))
                    Page.AddVariable("%su_sno" & StandardsCount & "%", Me.IsNull(Row.Item("serial_no")))
                    If ExternalDateExpires > InternalDateExpires Then
                        nr.Item("date_tested") = ExternalDateTested
                        nr.Item("date_expires") = ExternalDateExpires
                        nr.Item("test_no") = Row.Item("test_no")
                        nr.Item("uncertainty") = Row.Item("external_uncertainty_value")
                        Page.AddVariable("%su_tno" & StandardsCount & "%", Me.IsNull(Row.Item("test_no")))
                        Page.AddVariable("%su_exp" & StandardsCount & "%", ExternalDateExpires.ToString("MM/dd/yyyy"))
                        Page.AddVariable("%su_tst" & StandardsCount & "%", ExternalDateTested.ToString("MM/dd/yyyy"))
                        Page.AddVariable("%su_uc" & StandardsCount & "%", Me.IsNull(Row.Item("external_uncertainty_value")))
                    Else
                        nr.Item("date_tested") = InternalDateTested.ToString("MM/dd/yyyy")
                        nr.Item("date_expires") = InternalDateExpires.ToString("MM/dd/yyyy")
                        nr.Item("test_no") = "Internal"
                        nr.Item("uncertainty") = Row.Item("internal_uncertainty_value")
                        Page.AddVariable("%su_tno" & StandardsCount & "%", "Internal")
                        Page.AddVariable("%su_exp" & StandardsCount & "%", InternalDateExpires.ToString("MM/dd/yyyy"))
                        Page.AddVariable("%su_tst" & StandardsCount & "%", InternalDateTested.ToString("MM/dd/yyyy"))
                        Page.AddVariable("%su_uc" & StandardsCount & "%", Me.IsNull(Row.Item("internal_uncertainty_value")))
                    End If
                Else
                    Page.AddVariable("%su_ano" & StandardsCount & "%", "")
                    Page.AddVariable("%su_sno" & StandardsCount & "%", "")
                    Page.AddVariable("%su_tno" & StandardsCount & "%", "")
                    Page.AddVariable("%su_exp" & StandardsCount & "%", "")
                    Page.AddVariable("%su_tst" & StandardsCount & "%", "")
                    Page.AddVariable("%su_uc" & StandardsCount & "%", "")
                End If
                StandardsCount += 1
            Next
            ' Put table in template
            Dim Standards As GravityDocument.gElement = Page.GetTableBySource("standards_used")
            If Standards IsNot Nothing Then
                Standards.Table.Data = NewStandardsTable
            End If
        ElseIf Me.CustomerStandards.Length > 0 Then
            Page.AddVariable("%standards_used%", "Used customer stanadrds. " & Me.CustomerStandards)
        Else
            Page.AddVariable("%standards_used%", "None")
        End If
        ' Blank out any of these
        For i As Integer = StandardsCount - 1 To 20   ' Up to 21 standards used
            Page.AddVariable("%su_ano" & (i + 1) & "%", "")
            Page.AddVariable("%su_sno" & (i + 1) & "%", "")
            Page.AddVariable("%su_tno" & (i + 1) & "%", "")
            Page.AddVariable("%su_exp" & (i + 1) & "%", "")
            Page.AddVariable("%su_tst" & (i + 1) & "%", "")
            Page.AddVariable("%su_uc" & (i + 1) & "%", "")
        Next

        ' Techs
        Dim Users As String = Me.Database.Escape(so.AssignedTo) & ", "
        Dim strTechs As String = ""
        If Not Blank Then
            dt = Me.Database.GetAll("SELECT DISTINCT owner FROM schedule WHERE reference_id=" & so.OrderNo)
            For Each r In dt.Rows
                If so.AssignedTo.Length > 0 Then
                    If Not r.Item("owner") = so.AssignedTo Then
                        Users &= Me.Database.Escape(r.Item("owner")) & ", "
                    End If
                Else
                    Users &= Me.Database.Escape(r.Item("owner")) & ", "
                End If
            Next
            Users = Users.Substring(0, Users.Length - 2)
            Dim Techs As DataTable = Me.Database.GetAll("SELECT (last_name + ', ' + first_name) AS display_name FROM employee WHERE windows_user IN (" & Users & ")")
            For Each r In Techs.Rows
                strTechs &= r.Item("display_name") & "; "
            Next
            If strTechs.Length > 2 Then
                strTechs = strTechs.Substring(0, strTechs.Length - 2)
            End If
        End If
        Page.AddVariable("%techs%", strTechs)
        Page.AddVariable("%lead_tech%", Me.Technician)

        ' Additional Charges
        If Not Blank Then
            Page.AddVariable("%additional_total%", Format(Me.LineItemsTotal, "$#.00"))
            Page.AddVariable("%parts%", Format(Me.PartsTotal, "$#.00"))
            Page.AddVariable("%other%", Format(Me.TravelTotal + Me.LaborTotal, "$#.00"))
        Else
            Page.AddVariable("%additional_total%", "")
            Page.AddVariable("%parts%", "")
            Page.AddVariable("%other%", "")
        End If

        ' Equipment Condition
        If Not Blank Then
            Select Case Me.EquipmentCondition
                Case 1
                    Page.AddVariable("%eq_g%", "")
                    Page.AddVariable("%eq_f%", "")
                    Page.AddVariable("%eq_p%", "")
                    Page.AddVariable("%eq_vg%", "X")
                    Page.AddVariable("%eq_word%", "Very Good")
                    Page.AddVariable("%eq_abbr%", "VG")
                Case 2
                    Page.AddVariable("%eq_vg%", "")
                    Page.AddVariable("%eq_f%", "")
                    Page.AddVariable("%eq_p%", "")
                    Page.AddVariable("%eq_g%", "X")
                    Page.AddVariable("%eq_word%", "Good")
                    Page.AddVariable("%eq_abbr%", "G")
                Case 3
                    Page.AddVariable("%eq_vg%", "")
                    Page.AddVariable("%eq_g%", "")
                    Page.AddVariable("%eq_p%", "")
                    Page.AddVariable("%eq_f%", "X")
                    Page.AddVariable("%eq_word%", "Fair")
                    Page.AddVariable("%eq_abbr%", "F")
                Case 4
                    Page.AddVariable("%eq_vg%", "")
                    Page.AddVariable("%eq_g%", "")
                    Page.AddVariable("%eq_f%", "")
                    Page.AddVariable("%eq_p%", "X")
                    Page.AddVariable("%eq_word%", "Poor")
                    Page.AddVariable("%eq_abbr%", "P")
            End Select
        Else
            Page.AddVariable("%eq_vg%", "")
            Page.AddVariable("%eq_g%", "")
            Page.AddVariable("%eq_f%", "")
            Page.AddVariable("%eq_p%", "")
            Page.AddVariable("%eq_word%", "")
            Page.AddVariable("%eq_abbr%", "")
        End If

        ' Environmental Condition
        If Not Blank Then
            Page.AddVariable("%barometer%", Me.Barometer)
            Page.AddVariable("%temp%", Me.Temperature)
            Page.AddVariable("%humidity%", Me.Humidity)
            Select Case Me.EnvironmentalCondition
                Case 1
                    Page.AddVariable("%env_g%", "")
                    Page.AddVariable("%env_f%", "")
                    Page.AddVariable("%env_p%", "")
                    Page.AddVariable("%env_vg%", "X")
                    Page.AddVariable("%env_word%", "Very Good")
                    Page.AddVariable("%env_abbr%", "VG")
                Case 2
                    Page.AddVariable("%env_vg%", "")
                    Page.AddVariable("%env_f%", "")
                    Page.AddVariable("%env_p%", "")
                    Page.AddVariable("%env_g%", "X")
                    Page.AddVariable("%env_word%", "Good")
                    Page.AddVariable("%env_abbr%", "G")
                Case 3
                    Page.AddVariable("%env_vg%", "")
                    Page.AddVariable("%env_g%", "")
                    Page.AddVariable("%env_p%", "")
                    Page.AddVariable("%env_f%", "X")
                    Page.AddVariable("%env_word%", "Fair")
                    Page.AddVariable("%env_abbr%", "F")
                Case 4
                    Page.AddVariable("%env_vg%", "")
                    Page.AddVariable("%env_g%", "")
                    Page.AddVariable("%env_f%", "")
                    Page.AddVariable("%env_p%", "X")
                    Page.AddVariable("%env_word%", "Poor")
                    Page.AddVariable("%env_abbr%", "P")
            End Select
        Else
            Page.AddVariable("%env_vg%", "")
            Page.AddVariable("%env_g%", "")
            Page.AddVariable("%env_f%", "")
            Page.AddVariable("%env_p%", "")
            Page.AddVariable("%env_word%", "")
            Page.AddVariable("%env_abbr%", "")
            Page.AddVariable("%barometer%", "")
            Page.AddVariable("%temp%", "")
            Page.AddVariable("%humidity%", "")
        End If

        ' Uncertainty
        If Me.UncertaintyStatementId > 0 And Not Blank Then
            Page.AddVariable("%uc_statement%", Database.GetOne("SELECT statement FROM uncertainty_statement WHERE id=" & Me.UncertaintyStatementId))
        Else
            Page.AddVariable("%uc_statement%", "")
        End If
        Page.AddVariable("%uc_value%", "")
        Page.AddVariable("%uc_units%", "")

        ' Calibrated To
        If Not Blank Then
            Select Case Me.CalibrationProcedure
                Case 1
                    Page.AddVariable("%cal_to_man%", "")
                    Page.AddVariable("%cal_to_hb44%", "")
                    Page.AddVariable("%cal_to_man%", "")
                    Page.AddVariable("%cal_to_hb44%", "")
                    Page.AddVariable("%cal_to_cust%", "X")
                    Page.AddVariable("%cal_to%", "Customer")
                Case 2
                    Page.AddVariable("%cal_to_hb44%", "")
                    Page.AddVariable("%cal_to_cust%", "")
                    Page.AddVariable("%cal_to_man%", "X")
                    Page.AddVariable("%cal_to%", "Manufacturer")
                Case 3
                    Page.AddVariable("%cal_to_man%", "")
                    Page.AddVariable("%cal_to_cust%", "")
                    Page.AddVariable("%cal_to_hb44%", "X")
                    Page.AddVariable("%cal_to%", "Handbook 44")
                Case Else
                    Page.AddVariable("%cal_to_man%", "")
                    Page.AddVariable("%cal_to_hb44%", "")
                    Page.AddVariable("%cal_to_cust%", "")
                    Page.AddVariable("%cal_to%", "Unknown")
            End Select
        Else
            Page.AddVariable("%cal_to_man%", "")
            Page.AddVariable("%cal_to_hb44%", "")
            Page.AddVariable("%cal_to_cust%", "")
            Page.AddVariable("%cal_to%", "--")
        End If

        ' NEXT CAL DATE
        Dim NextCal As Date
        Dim PrevCal As Date = Today
        ' Completed?
        If so.DateCompleted = Nothing Then
            ' Calibrated?
            If Not Me.Calibrated Then
                If Me.EquipmentId > 0 Then
                    Dim Equip As New cEquipment(Me.Database)
                    Equip.Open(Me.EquipmentId)
                    PrevCal = Equip.DateLastCalibrated
                Else
                    PrevCal = Nothing
                End If
            Else
                PrevCal = Nothing
            End If
        Else
            ' Calibrated?
            ' If it was not calibrated then we must get the previous cal date, if it was never
            ' calibrated previously (at least not on our records) make prev cal nothing.
            ' If it WAS calibrated on this trip then the completed date of this service order
            ' is the previous cal date
            If Not Me.Calibrated Then
                Try
                    If Me.EquipmentId > 0 Then
                        Dim Equip As New cEquipment(Me.Database)
                        Equip.Open(Me.EquipmentId)
                        PrevCal = Equip.DateLastCalibrated
                    Else
                        PrevCal = Nothing
                    End If
                Catch
                    PrevCal = Nothing
                End Try
            Else
                PrevCal = so.DateCompleted
            End If
        End If
        ' Get next due
        If Me.DateNextDue = Nothing Then
            NextCal = Nothing       ' If it has never been calibrated, we can't calculate next due
        Else
            NextCal = Me.DateNextDue
        End If
        ' Write it
        If NextCal = Nothing Then
            Page.AddVariable("%next_cal_date%", "--")
        Else
            Page.AddVariable("%next_cal_date%", Format(NextCal, "MM/dd/yyyy"))
        End If
        If PrevCal = Nothing Then
            Page.AddVariable("%calibration_date%", "--")
        Else
            Page.AddVariable("%calibration_date%", Format(PrevCal, "MM/dd/yyyy"))
        End If
        ' Frequency
        ' On a cal agreement
        Dim Settings As New cSettings(Me.Database)
        If Not so.CalAgreement = 0 Then
            Dim Ca As DataRow = Database.GetRow("SELECT frequency_months FROM cal_agreement WHERE id=" & so.CalAgreement)
            If Ca IsNot Nothing Then
                Page.AddVariable("%cal_interval%", Ca.Item("frequency_months"))
            Else
                Page.AddVariable("%cal_interval%", Math.Round(Settings.GetValue("Next Due Days") / 30))
            End If
        Else
            Page.AddVariable("%cal_interval%", Math.Round(Settings.GetValue("Next Due Days") / 30))
        End If

        ' Tests
        Dim CalTests As DataTable = Me.CalibrationTests
        Dim dblRepeatability As Double = Me.Repeatability
        If dblRepeatability < 0 Then
            Page.AddVariable("%repeatability%", "--")
        Else
            Page.AddVariable("%repeatability%", dblRepeatability)
        End If
        ' Build new test data table
        Dim TestTable As New DataTable
        TestTable.Columns.Add("instrument")
        TestTable.Columns.Add("test_name")
        TestTable.Columns.Add("applied")
        TestTable.Columns.Add("initial")
        TestTable.Columns.Add("final")
        TestTable.Columns.Add("note")
        TestTable.Columns.Add("found_pass_yn")
        TestTable.Columns.Add("found_pass_pf")
        TestTable.Columns.Add("found_pass_yesno")
        TestTable.Columns.Add("found_pass_x")
        TestTable.Columns.Add("left_pass_yn")
        TestTable.Columns.Add("left_pass_pf")
        TestTable.Columns.Add("left_pass_yesno")
        TestTable.Columns.Add("left_pass_x")
        TestTable.Columns.Add("initial_error")
        TestTable.Columns.Add("initial_error_divisions")
        TestTable.Columns.Add("final_error")
        TestTable.Columns.Add("final_error_divisions")
        TestTable.Columns.Add("tolerance_value")
        TestTable.Columns.Add("repeatability")
        TestTable.Columns.Add("uncertainty")
        ' Loop through and make new table
        For Each Row As DataRow In CalTests.Rows
            Dim nr As DataRow = TestTable.NewRow
            nr.Item("instrument") = Row.Item("instrument")
            nr.Item("test_name") = Row.Item("test_name")
            nr.Item("applied") = Row.Item("applied")
            nr.Item("initial") = Row.Item("initial")
            nr.Item("final") = Row.Item("final")
            nr.Item("note") = Row.Item("note")
            nr.Item("tolerance_value") = Row.Item("tolerance_value")
            If Row.Item("as_found_pass") IsNot DBNull.Value Then
                nr.Item("found_pass_yn") = IIf(Row.Item("as_found_pass"), "Y", "N")
                nr.Item("found_pass_pf") = IIf(Row.Item("as_found_pass"), "P", "F")
                nr.Item("found_pass_yesno") = IIf(Row.Item("as_found_pass"), "Yes", "No")
                nr.Item("found_pass_x") = IIf(Row.Item("as_found_pass"), "X", "")
            Else
                nr.Item("found_pass_yn") = "--"
                nr.Item("found_pass_pf") = "--"
                nr.Item("found_pass_yesno") = "--"
                nr.Item("found_pass_x") = "--"
            End If
            If Row.Item("pass") IsNot DBNull.Value Then
                nr.Item("left_pass_yn") = IIf(Row.Item("pass"), "Y", "N")
                nr.Item("left_pass_pf") = IIf(Row.Item("pass"), "P", "F")
                nr.Item("left_pass_yesno") = IIf(Row.Item("pass"), "Yes", "No")
                nr.Item("left_pass_x") = IIf(Row.Item("pass"), "X", "")
            Else
                nr.Item("left_pass_yn") = "--"
                nr.Item("left_pass_pf") = "--"
                nr.Item("left_pass_yesno") = "--"
                nr.Item("left_pass_x") = "--"
            End If
            ' Calculate Error
            Dim DecimalPlaces As Integer = 3
            Try
                DecimalPlaces = Math.Max(Me.CountDecimals(Me.Equipment.CountBy), Me.CountDecimals(Row.Item("initial")))
            Catch ex As Exception
                ' Ignore
            End Try
            Try
                Dim Applied As Double = Me.FloatVal(Row.Item("applied"))
                Dim Initial As Double = Me.FloatVal(Row.Item("initial"))
                Dim Final As Double = Me.FloatVal(Row.Item("final"))
                Dim CountBy As Double = Me.FloatVal(Me.Equipment.CountBy)
                Dim InitErr As Double = (Applied - Initial) * -1
                Dim FinalErr As Double = (Applied - Final) * -1
                If (Applied <> Nothing And Initial <> Nothing) Or (Applied = 0 And Initial = 0) Then
                    nr.Item("initial_error") = Me.MyRounding(InitErr, DecimalPlaces)
                    If CountBy <> Nothing And InitErr <> 0 Then
                        nr.Item("initial_error_divisions") = Math.Round(InitErr / CountBy)
                    ElseIf InitErr = 0 Then
                        nr.Item("initial_error_divisions") = 0
                    Else
                        nr.Item("initial_error_divisions") = "--"
                    End If
                Else
                    nr.Item("initial_error") = "--"
                    nr.Item("initial_error_divisions") = "--"
                End If
                If (Applied <> Nothing And Final <> Nothing) Or (Applied = 0 And Final = 0) Then
                    nr.Item("final_error") = Me.MyRounding(FinalErr, DecimalPlaces)
                    If CountBy <> Nothing And FinalErr <> 0 Then
                        nr.Item("final_error_divisions") = Math.Round(FinalErr / CountBy)
                    ElseIf FinalErr = 0 Then
                        nr.Item("final_error_divisions") = 0
                    Else
                        nr.Item("final_error_divisions") = "--"
                    End If
                Else
                    nr.Item("final_error") = "--"
                    nr.Item("final_error_divisions") = "--"
                End If
            Catch
                nr.Item("initial_error") = "--"
                nr.Item("initial_error_divisions") = "--"
                nr.Item("final_error") = "--"
                nr.Item("final_error_divisions") = "--"
            End Try
            ' Repeatability/Uncertainty
            If dblRepeatability >= 0 Then
                'MsgBox(dblRepeatability)
                Dim strRepeatability As String = Me.MyRounding(dblRepeatability, DecimalPlaces + 2)
                nr.Item("repeatability") = strRepeatability
                nr.Item("uncertainty") = Me.Uncertainty(Row, strRepeatability, DecimalPlaces + 2)
            Else
                nr.Item("repeatability") = "--"
                nr.Item("uncertainty") = "--"
            End If
            ' Add row
            TestTable.Rows.Add(nr)
        Next
        ' Add table
        Dim Tests As GravityDocument.gElement = Page.GetTableBySource("tests")
        If Tests IsNot Nothing Then
            Tests.Table.Data = TestTable
        End If
        ' Add non-table variables
        For i As Integer = 0 To 20   ' Up to 21 test points
            If TestTable.Rows.Count > i Then
                Page.AddVariable("%inst" & (i + 1) & "%", Me.IsNull(TestTable.Rows(i).Item("instrument")))
                Page.AddVariable("%tn" & (i + 1) & "%", Me.IsNull(TestTable.Rows(i).Item("test_name")))
                Page.AddVariable("%app" & (i + 1) & "%", Me.IsNull(TestTable.Rows(i).Item("applied")))
                Page.AddVariable("%init" & (i + 1) & "%", Me.IsNull(TestTable.Rows(i).Item("initial")))
                Page.AddVariable("%afp" & (i + 1) & "%", Me.IsNull(TestTable.Rows(i).Item("found_pass_yn")))
                Page.AddVariable("%final" & (i + 1) & "%", Me.IsNull(TestTable.Rows(i).Item("final")))
                Page.AddVariable("%alp" & (i + 1) & "%", Me.IsNull(TestTable.Rows(i).Item("left_pass_yn")))
                Page.AddVariable("%tnote" & (i + 1) & "%", Me.IsNull(TestTable.Rows(i).Item("note")))
                Page.AddVariable("%ierr" & (i + 1) & "%", Me.IsNull(TestTable.Rows(i).Item("initial_error")))
                Page.AddVariable("%ierrd" & (i + 1) & "%", Me.IsNull(TestTable.Rows(i).Item("initial_error_divisions")))
                Page.AddVariable("%ferr" & (i + 1) & "%", Me.IsNull(TestTable.Rows(i).Item("final_error")))
                Page.AddVariable("%ferrd" & (i + 1) & "%", Me.IsNull(TestTable.Rows(i).Item("final_error_divisions")))
                Page.AddVariable("%tol" & (i + 1) & "%", Me.IsNull(TestTable.Rows(i).Item("tolerance_value")))
            Else
                Page.AddVariable("%inst" & (i + 1) & "%", "")
                Page.AddVariable("%tn" & (i + 1) & "%", "")
                Page.AddVariable("%app" & (i + 1) & "%", "")
                Page.AddVariable("%init" & (i + 1) & "%", "")
                Page.AddVariable("%afp" & (i + 1) & "%", "")
                Page.AddVariable("%final" & (i + 1) & "%", "")
                Page.AddVariable("%alp" & (i + 1) & "%", "")
                Page.AddVariable("%tnote" & (i + 1) & "%", "")
                Page.AddVariable("%ierr" & (i + 1) & "%", "")
                Page.AddVariable("%ierrd" & (i + 1) & "%", "")
                Page.AddVariable("%ferr" & (i + 1) & "%", "")
                Page.AddVariable("%ferrd" & (i + 1) & "%", "")
                Page.AddVariable("%tol" & (i + 1) & "%", "")
            End If
        Next

        Dim Charges As GravityDocument.gElement = Page.GetTableBySource("additional_charges")
        If Charges IsNot Nothing Then
            Charges.Table.Data = Me.WorkOrderItems
        End If
        Return Doc
    End Function

    Private Function FloatVal(ByVal str As String) As Double
        If Microsoft.VisualBasic.IsNumeric(str) Then
            Return str
        ElseIf ("_" & str).IndexOfAny("0123456789") > 0 Then
            Dim NumStarted As Boolean = False
            Dim out As String = ""
            For i As Integer = 0 To str.Length - 1
                Dim c As Char = str.Substring(i, 1)
                If Microsoft.VisualBasic.IsNumeric(c) Then
                    NumStarted = True
                    out &= c
                ElseIf NumStarted And c = "." Then
                    out &= c
                ElseIf NumStarted Then
                    Exit For
                End If
            Next
            If out.EndsWith(".") Then
                out = out.Substring(0, out.Length - 1)
            End If
            Return out
        Else
            Return Nothing
        End If
    End Function

    Private Function CountDecimals(ByVal str As String) As Integer
        Dim n As Integer = 0
        Try
            If str.Contains(".") Then
                Dim Pos As Integer = str.Trim.LastIndexOf(".")
                Return str.Trim.Length - (Pos + 1)
            End If
        Catch
            ' Ignore
        End Try
        Return n
    End Function

    Private Function MyRounding(ByVal num As Decimal, Optional ByVal decimals As Integer = 0) As String
        num = Math.Round(num, decimals)
        Dim CurrentDecimals As Integer = Me.CountDecimals(num)
        If decimals > 0 And CurrentDecimals < decimals Then
            Dim out As String = num
            If Not num.ToString.Contains(".") Then
                out &= "."
            End If
            For i As Integer = CurrentDecimals To decimals - 1
                out &= "0"
            Next
            Return out
        Else
            Return num.ToString
        End If
    End Function

    Public Function GetCalibrationTests() As DataTable
        Dim Sql As String = "SELECT t.[id], work_order_id, instrument, test_name, applied, initial, final,"
        Sql &= " note, pass, as_found_pass, tolerance_value, t.date_last_updated, ctt.is_repeatability"
        Sql &= " FROM calibration_test t"
        Sql &= " LEFT OUTER JOIN calibration_test_type ctt ON t.test_name=ctt.name"
        Sql &= " WHERE work_order_id=" & Me._Id
        Return Me.Database.GetAll(Sql)
    End Function

    Private Function StandardDeviation(ByVal data As ArrayList) As Double
        ' Get Sum
        Dim Sum As Double = 0
        For i As Integer = 0 To data.Count - 1
            Try
                Sum += data(i)
            Catch ex As Exception
                ' Ignore
            End Try
        Next
        ' Get mean
        Dim Mean As Double = Sum / data.Count
        ' Get N
        Dim N As Double = 0
        For i As Integer = 0 To data.Count - 1
            Try
                N += (data(i) - Mean) ^ 2
            Catch ex As Exception
                ' Ignore
            End Try
        Next
        ' Get X
        Dim X As Double = N / (data.Count - 1)
        ' Return standard deviation
        Try
            Return Math.Sqrt(X)
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Public Function Uncertainty(ByVal Row As DataRow, ByVal dblRepeatability As String, Optional ByVal DecimalCount As Integer = 4) As String
        Dim Out As Double = 0
        Try
            If Me.Equipment.TolerenaceUncertaintyFormula.Length = 0 Then
                Return "No Formula"
            ElseIf dblRepeatability < 0 Then
                Return "N/A"
            Else
                'MsgBox(dblRepeatability)
                Dim Formula As New MyCore.mcCalc
                Formula.AddVariable("repeatability", dblRepeatability)
                Formula.AddVariable("applied", Row.Item("applied"))
                Formula.AddVariable("initial", Row.Item("initial"))
                Formula.AddVariable("final", Row.Item("final"))
                Formula.AddVariable("graduation", Me.Equipment.CountBy)
                Out = Formula.Eval(Me.Equipment.TolerenaceUncertaintyFormula)
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
        If Out > 0 Then
            Return Me.MyRounding(Out, DecimalCount)
        Else
            Return "??"
        End If
    End Function

    Public Function Repeatability() As Double
        Dim StDev As Double = -1
        Dim Sum As Double = 0
        Dim ReptTests As New ArrayList
        If Me.CalibrationTests IsNot Nothing Then
            For Each Row As DataRow In Me.CalibrationTests.Rows
                If IIf(Row.Item("is_repeatability") Is DBNull.Value, False, Row.Item("is_repeatability")) Then
                    Try
                        ' Will fail if not a number
                        Sum += CDbl(Row.Item("final"))
                        ReptTests.Add(Row.Item("final"))
                    Catch ex As Exception
                        ' Ignore
                    End Try
                End If
            Next
            If ReptTests.Count > 0 Then
                StDev = Me.StandardDeviation(ReptTests)
            End If
        End If
        Return StDev
    End Function

    Public Function NextDue() As Date
        ' *** Later need to calculate this
        Return Me.CalibrationTimestamp.AddDays(90)
    End Function

    Public Function GetStandardsUsed() As DataTable
        Dim Sql As String = ""
        Sql &= "SELECT e.*,"
        Sql &= " c.test_no, c.date_tested, c.date_expires, c.id AS cert_id, s.name AS location_name"
        Sql &= " FROM standards_equipment e"
        Sql &= " LEFT OUTER JOIN station s ON e.station_id=s.id"
        Sql &= " LEFT OUTER JOIN (SELECT stc.asset_no, c.* FROM standards_to_cert stc LEFT JOIN standards_certification c ON stc.test_no=c.test_no) c"
        Sql &= " ON e.asset_no=c.asset_no"
        Sql &= " WHERE e.asset_no IN (SELECT asset_no FROM standards_used WHERE work_order_id=" & Me.Id & ")"
        Sql &= " ORDER BY e.asset_no, date_expires DESC"
        Dim Table As DataTable = Me.Database.GetAll(Sql)
        Dim StandardsUsed As DataTable = Table.Clone
        Dim AssetHash As New Hashtable
        For Each Row As DataRow In Table.Rows
            If Not AssetHash.ContainsKey(Row.Item("asset_no")) Then
                Dim nr As DataRow = StandardsUsed.NewRow
                For i As Integer = 0 To Row.ItemArray.Length - 1
                    nr.Item(i) = Row.Item(i)
                Next
                StandardsUsed.Rows.Add(nr)
                AssetHash.Add(Row.Item("asset_no"), "")
            End If
        Next
        Return StandardsUsed
    End Function



End Class
