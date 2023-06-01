Imports MyCore.Data

Public Class cEquipment

    Dim _Id As Integer = 0
    Public PartNo As String = ""
    Public Manufacturer As String = ""
    Public Model As String = ""
    Public Serial As String = ""
    Public ShipTo As String = ""
    Public Status As String = "INV"
    Public Station As Integer = 0
    Public Capacity As String = ""
    Public CapacityUnits As Integer = 0
    Public Category As Integer = 0
    Public CountBy As String = ""
    Public CountByUnits As Integer = 0
    Public AccuracyClass As Integer = 0
    Public Options As String = ""
    Public AttachedTo As Integer = 0
    Public DualRemote As Boolean = False
    Public AssetNo As String = ""
    Public CustomerDescription As String = ""
    Public Inactive As Boolean = False
    Public Cost As Double = 0
    Public ListPrice As Double = 0
    Public Freight As Double = 0
    Public Office As String = ""
    Public Bought As Date = Nothing
    Public Aquired As Date = Nothing
    Public Condition As String = ""
    Public OurPO As String = ""
    Public InvoiceNo As String = ""
    Public Salesman As String = ""
    Public GMAExpire As Date = Nothing
    Public PartsExpire As Date = Nothing
    Public LaborExpire As Date = Nothing
    Public PreviousCal As Date = Nothing
    Public NextCal As Date = Nothing
    Public CalAgreement As Integer = 0
    Public CalProcedure As Integer = 0
    Public Tolerenace As String = ""
    Public TolerenaceName As String = ""
    Public TolerenaceUncertaintyFormula As String = ""
    Public CalTemplateId As Integer = 0
    Public CertTemplateId As Integer = 0

    Public LastUpdatedBy As String = ""
    Public CreatedBy As String = ""

    Dim OldStatus As String = ""
    Dim OldShipTo As String = ""
    Dim OldStation As Integer = 0

    Dim _Condition As DataTable
    Dim _Category As DataTable
    Dim _Status As DataTable
    Dim _Station As DataTable
    Dim _Manufacturer As DataTable
    Dim _Units As DataTable
    Dim _Office As DataTable
    Dim _AccuracyClasses As DataTable
    Dim _ErrorMsg As String = ""

    Dim Database As New MyCore.Data.EasySql("MsSql")

    Public Event Reload()
    Public Event Saved(ByVal Equipment As cEquipment)

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

    Public ReadOnly Property Conditions() As DataTable
        Get
            Return Me._Condition
        End Get
    End Property

    Public ReadOnly Property Categories() As DataTable
        Get
            Return Me._Category
        End Get
    End Property

    Public ReadOnly Property Statuses() As DataTable
        Get
            Return Me._Status
        End Get
    End Property

    Public ReadOnly Property Stations() As DataTable
        Get
            Return Me._Station
        End Get
    End Property

    Public ReadOnly Property Manufacturers() As DataTable
        Get
            Return Me._Manufacturer
        End Get
    End Property

    Public ReadOnly Property Units() As DataTable
        Get
            Return Me._Units
        End Get
    End Property

    Public ReadOnly Property AccuracyClasses() As DataTable
        Get
            Return Me._AccuracyClasses
        End Get
    End Property

    Public ReadOnly Property CountByUnitsName() As String
        Get
            If Me.CountByUnits > 0 Then
                Return Me.Database.GetOne("SELECT name FROM units_of_measure WHERE id=" & Me.CountByUnits)
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property CapacityUnitsName() As String
        Get
            If Me.CapacityUnits > 0 Then
                Return Me.Database.GetOne("SELECT name FROM units_of_measure WHERE id=" & Me.CapacityUnits)
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property DateLastCalibrated() As Date
        Get
            Dim r As DataRow = Me.Database.GetRow("SELECT TOP 1 so.date_completed FROM work_order wo LEFT OUTER JOIN service_order so ON wo.service_order_id=so.id WHERE calibrated=1 AND equipment_id=" & Me.Id & " ORDER BY so.date_completed DESC)")
            If Me.Database.LastQuery.RowsReturned > 0 Then
                If r.Item("date_completed") IsNot DBNull.Value Then
                    Return r.Item("date_completed")
                End If
            End If
            Return Nothing
        End Get
    End Property


    Public Sub New(ByRef db As MyCore.Data.EasySql)
        Me.Database = db
        Me.PopulateOffice()
        Me.PopulateCategory()
        Me.PopulateManufacturer()
        Me.PopulateStation()
        Me.PopulateStatus()
        Me.PopulateUnits()
        Me.PopulateTolerances()
        Me.PopulateCondition()
    End Sub

    Public Sub New(ByVal Id As Integer)
        Me._Id = Id
        Me.PopulateOffice()
        Me.PopulateCategory()
        Me.PopulateManufacturer()
        Me.PopulateStation()
        Me.PopulateStatus()
        Me.PopulateUnits()
        Me.PopulateTolerances()
        Me.PopulateCondition()
    End Sub

    Private Sub PopulateOffice()
        Me._Office = Me.Database.GetAll("SELECT id, number, name, sort FROM office ORDER BY sort, name")
    End Sub

    Private Sub PopulateCategory()
        Me._Category = Me.Database.GetAll("SELECT id, name, sort FROM item_category ORDER BY sort, name")
    End Sub

    Private Sub PopulateCondition()
        Me._Condition = Me.Database.GetAll("SELECT id, name, sort FROM condition ORDER BY sort, name")
    End Sub

    Private Sub PopulateManufacturer()
        Me._Manufacturer = Me.Database.GetAll("SELECT id, name, sort FROM manufacturer ORDER BY sort, name")
    End Sub

    Private Sub PopulateStation()
        Me._Station = Me.Database.GetAll("SELECT id, name, sort FROM station ORDER BY sort, name")
    End Sub

    Private Sub PopulateStatus()
        Me._Status = Me.Database.GetAll("SELECT id, code, name, sort FROM equipment_status ORDER BY sort, name")
    End Sub

    Private Sub PopulateUnits()
        Me._Units = Me.Database.GetAll("SELECT id, name, sort FROM units_of_measure ORDER BY sort, name")
    End Sub

    Private Sub PopulateTolerances()
        Me._AccuracyClasses = Me.Database.GetAll("SELECT id, name, sort FROM tolerance ORDER BY sort, name")
    End Sub

    Public Function CalTemplates() As DataTable
        Return Me.Database.GetAll("SELECT id, name FROM calibration_template ORDER BY name")
    End Function

    Public Sub Open(ByVal Id As Integer)
        Dim Sql As String = "SELECT"
        Sql &= " name = (equip.dep_manuf + ' ' + equip.dep_mod + ' (' + equip.dep_ser + ')'),"
        Sql &= " equip.*,"
        Sql &= " tolerance.name AS tolerance_name, tolerance.uncertainty_formula,"
        Sql &= " ADDRESS.cst_name, ADDRESS.cst_city, ADDRESS.cst_state, "
        Sql &= " ISNULL(ADDRESS.cst_zip, '') AS cst_zip, "
        Sql &= " indicator_name =  indicator.dep_manuf + ' ' + indicator.dep_mod  + ' (' + indicator.dep_ser + ')',"
        Sql &= " date_last_cal = (SELECT TOP 1 so.date_completed"
        Sql &= " FROM work_order wo RIGHT JOIN service_order so ON wo.service_order_id=so.id"
        Sql &= " WHERE calibrated=1 AND equipment_id=" & Id & " ORDER BY so.date_completed DESC)"
        Sql &= " FROM DEPREC equip"
        Sql &= " LEFT OUTER JOIN ADDRESS ON equip.dep_loc = ADDRESS.cst_no "
        Sql &= " LEFT OUTER JOIN DEPREC indicator ON equip.dep_indicator=indicator.dep_id"
        Sql &= " LEFT OUTER JOIN tolerance ON equip.tolerance_id=tolerance.id"
        Sql &= " WHERE equip.dep_id=" & Id
        Dim Row As DataRow = Me.Database.GetRow(Sql)
        If Me.Database.LastQuery.RowsReturned = 1 Then
            Me._Id = Row.Item("dep_id")
            If Row.Item("dep_date") Is DBNull.Value Then
                Me.Aquired = Nothing
            Else
                Me.Aquired = Row.Item("dep_date")
            End If
            Me.AssetNo = Me.IsNull(Row.Item("dep_assno"), "")
            If Row.Item("dep_indicator") Is DBNull.Value Then
                Me.AttachedTo = Nothing
            Else
                Me.AttachedTo = Row.Item("dep_indicator")
            End If
            If Row.Item("dep_dbot") Is DBNull.Value Then
                Me.Bought = Nothing
            Else
                Me.Bought = Row.Item("dep_dbot")
            End If
            Me.CalAgreement = 0
            Me.Capacity = Me.IsNull(Row.Item("dep_cap"), "")
            Me.CapacityUnits = Row.Item("capacity_units")
            Me.Category = Me.IsNull(Row.Item("dep_type"), 3)
            Me.Condition = Me.IsNull(Row.Item("dep_new"), Nothing)
            Me.Cost = Me.IsNull(Row.Item("dep_cost"), 0)
            Me.CountBy = Me.IsNull(Row.Item("dep_CountBy"), 0)
            Me.CountByUnits = Row.Item("count_by_units")
            Me.CustomerDescription = Row.Item("customer_description")
            Me.DualRemote = Row.Item("dual_remote")
            Me.Freight = Me.IsNull(Row.Item("dep_frt"), 0)
            If Row.Item("dep_gmad") Is DBNull.Value Then
                Me.GMAExpire = Nothing
            Else
                Me.GMAExpire = Row.Item("dep_gmad")
            End If
            Me.Inactive = Row.Item("inactive")
            Me.InvoiceNo = Me.IsNull(Row.Item("dep_inv"), "")
            If Row.Item("labor_expiration") Is DBNull.Value Then
                Me.LaborExpire = Nothing
            Else
                Me.LaborExpire = Row.Item("labor_expiration")
            End If
            Me.ListPrice = Me.IsNull(Row.Item("dep_price"), 0)
            Me.Manufacturer = Me.IsNull(Row.Item("dep_manuf"), "")
            Me.Model = Me.IsNull(Row.Item("dep_mod"), "")
            Me.NextCal = Nothing
            Me.Office = Me.IsNull(Row.Item("dep_off"), Nothing)
            Me.Options = Me.IsNull(Row.Item("dep_opt"), "")
            Me.OurPO = Me.IsNull(Row.Item("dep_ourpo"), "")
            Me.PartNo = Row.Item("part_no")
            If Row.Item("parts_expiration") Is DBNull.Value Then
                Me.PartsExpire = Nothing
            Else
                Me.PartsExpire = Row.Item("parts_expiration")
            End If
            If Row.Item("date_last_cal") Is DBNull.Value Then
                Me.PreviousCal = Nothing
            Else
                Me.PreviousCal = Row.Item("date_last_cal")
            End If
            Me.Salesman = Me.IsNull(Row.Item("dep_slsmn"), "")
            Me.Serial = Me.IsNull(Row.Item("dep_ser"), "")
            Me.ShipTo = Me.IsNull(Row.Item("dep_loc"), Nothing)
            Me.Station = Row.Item("station_id")
            Me.Status = Me.IsNull(Row.Item("dep_stat"), "")
            Me.AccuracyClass = Row.Item("tolerance_id")
            Me.Tolerenace = Me.IsNull(Row.Item("tolerance"), "")
            Me.TolerenaceName = Me.IsNull(Row.Item("tolerance_name"), "")
            Me.TolerenaceUncertaintyFormula = Me.IsNull(Row.Item("uncertainty_formula"), "")
            Me.CalProcedure = Row.Item("procedure_id")
            Me.CalTemplateId = Row.Item("cal_template_id")
            Me.CertTemplateId = Row.Item("cert_template_id")
            Me.LastUpdatedBy = Me.IsNull(Row.Item("dep_user"), "")
            Me.CreatedBy = Me.IsNull(Row.Item("dep_user"), "")
            RaiseEvent Reload()
            ' Save previous status
            Me.OldStatus = Me.Status
            Me.OldStation = Me.Station
            Me.OldShipTo = Me.ShipTo
            If Row.Table.Columns.Contains("tolerance_id") Then

            End If
        Else
            If Me.Database.LastQuery.Successful Then
                Throw New Exception("Could not open equipment id# " & Id)
            Else
                Throw New Exception(Me.Database.LastQuery.ErrorMsg)
            End If
        End If
    End Sub

    Public Sub Save(Optional ByVal Check As Boolean = True)
        If Check Then
            ' Check if complete
            If Not Me.IsComplete Then
                Throw New Exception("Form not completed. " & Me._ErrorMsg)
                Exit Sub
            End If
            ' Check for duplicate
            If Me.CheckForDuplicate Then
                Throw New Exception("Found duplicate equipment with same manufacturer, model, and serial number.")
                Exit Sub
            End If
        End If
        ' Continue
        Dim Sql As String = ""
        If Me._Id > 0 Then
            Sql = "UPDATE DEPREC SET"
            Sql &= " dep_date=@date_aquired, dep_manuf=@manufacturer, dep_mod=@model,"
            Sql &= " dep_ser=@serial, dep_cap=@capacity, dep_cost=" & Me.Database.ToCurrency("@cost") & ","
            Sql &= " dep_price=" & Me.Database.ToCurrency("@price") & ", dep_off=@office, dep_loc=@location,"
            Sql &= " dep_inv=@invoice, dep_ourpo=@our_po, dep_gmad=@gma_expiration,"
            Sql &= " dep_opt=@options, dep_dbot=@date_bought, dep_stat=@status,"
            Sql &= " dep_slsmn=@salesman, dep_new=@new, dep_assno=@asset_no,"
            Sql &= " dep_frt=" & Me.Database.ToCurrency("@freight") & ", dep_indicator=@indicator,"
            Sql &= " dep_type=@type, dep_CountBy=@count_by, labor_expiration=@labor_expiration,"
            Sql &= " parts_expiration=@parts_expiration, dep_chngt=" & Me.Database.Timestamp & ", dep_user=@updated_by,"
            Sql &= " customer_description=@customer_description, part_no=@part_no,"
            Sql &= " capacity_units=@capacity_units, count_by_units=@count_by_units,"
            Sql &= " tolerance_id=@tolerance_id, dual_remote=@dual_remote,"
            Sql &= " inactive=@inactive, station_id=@station, tolerance=@tolerance,"
            Sql &= " procedure_id=@procedure_id, cal_template_id=@cal_template_id,"
            Sql &= " cert_template_id=@cert_template_id "
            Sql &= " WHERE dep_id=@equipment_id"
        Else
            Sql = "INSERT INTO DEPREC"
            Sql &= " (dep_date, dep_manuf, dep_mod, dep_ser, dep_cap, dep_cost, dep_price, dep_off, dep_loc, dep_inv,"
            Sql &= " dep_ourpo, dep_gmad, dep_opt, dep_dbot, dep_stat, dep_slsmn, dep_new, dep_assno,"
            Sql &= " dep_frt, dep_indicator, dep_type, dep_CountBy, labor_expiration, parts_expiration, date_last_updated,"
            Sql &= " last_updated_by, customer_description, station_id, capacity_units, count_by_units, tolerance_id,"
            Sql &= " dual_remote, inactive, part_no, dep_chngt, dep_user,"
            Sql &= " cal_template_id, procedure_id, cert_template_id, tolerance"
            Sql &= " ) VALUES ("
            Sql &= " @date_aquired, @manufacturer, @model, @serial, @capacity, " & Me.Database.ToCurrency("@cost") & ", "
            Sql &= Me.Database.ToCurrency("@price") & ", @office, @location, @invoice, @our_po, @gma_expiration, @options,"
            Sql &= " @date_bought, @status, @salesman, @new, @asset_no, " & Me.Database.ToCurrency("@freight") & ", @indicator, "
            Sql &= " @type, @count_by,"
            Sql &= " @labor_expiration, @parts_expiration, " & Me.Database.Timestamp & ", @updated_by, @customer_description,"
            Sql &= " @station, @capacity_units, @count_by_units, @tolerance_id, @dual_remote, @inactive,"
            Sql &= " @part_no, " & Me.Database.Timestamp & ", SUBSTRING(@updated_by, 0, 6),"
            Sql &= " @cal_template_id, @procedure_id, @cert_template_id, @tolerance"
            Sql &= " )"
        End If
        Sql = Sql.Replace("@equipment_id", Me._Id)
        Sql = Sql.Replace("@date_aquired", Me.Database.Escape(Me.IsNothing(Me.Aquired, DBNull.Value)))
        Sql = Sql.Replace("@manufacturer", Me.Database.Escape(Me.Manufacturer))
        Sql = Sql.Replace("@model", Me.Database.Escape(Me.Model))
        Sql = Sql.Replace("@serial", Me.Database.Escape(Me.Serial))
        Sql = Sql.Replace("@capacity_units", Me.CapacityUnits)
        Sql = Sql.Replace("@count_by_units", Me.CountByUnits)
        Sql = Sql.Replace("@capacity", Me.Database.Escape(Me.Capacity))
        Sql = Sql.Replace("@cost", Me.Database.Escape(Me.Cost))
        Sql = Sql.Replace("@price", Me.Database.Escape(Me.ListPrice))
        Sql = Sql.Replace("@office", Me.Database.Escape(Me.Office))
        Sql = Sql.Replace("@location", Me.Database.Escape(Me.ShipTo))
        Sql = Sql.Replace("@invoice", Me.Database.Escape(Me.InvoiceNo))
        Sql = Sql.Replace("@our_po", Me.Database.Escape(Me.OurPO))
        Sql = Sql.Replace("@gma_expiration", Me.Database.Escape(Me.IsNothing(Me.GMAExpire, DBNull.Value)))
        Sql = Sql.Replace("@options", Me.Database.Escape(Me.Options))
        Sql = Sql.Replace("@date_bought", Me.Database.Escape(Me.IsNothing(Me.Bought, DBNull.Value)))
        Sql = Sql.Replace("@status", Me.Database.Escape(Me.Status))
        Sql = Sql.Replace("@salesman", Me.Database.Escape(Me.Salesman))
        Sql = Sql.Replace("@new", Me.Database.Escape(Me.Condition))
        Sql = Sql.Replace("@asset_no", Me.Database.Escape(Me.AssetNo))
        Sql = Sql.Replace("@freight", Me.Database.Escape(Me.Freight))
        Sql = Sql.Replace("@indicator", Me.Database.Escape(Me.AttachedTo))
        Sql = Sql.Replace("@type", Me.Category)
        Sql = Sql.Replace("@count_by", Me.Database.Escape(Me.CountBy))
        Sql = Sql.Replace("@labor_expiration", Me.Database.Escape(Me.IsNothing(Me.LaborExpire, DBNull.Value)))
        Sql = Sql.Replace("@parts_expiration", Me.Database.Escape(Me.IsNothing(Me.PartsExpire, DBNull.Value)))
        Sql = Sql.Replace("@updated_date", Me.Database.Escape(Now))
        Sql = Sql.Replace("@updated_by", Me.Database.Escape(Me.LastUpdatedBy))
        Sql = Sql.Replace("@customer_description", Me.Database.Escape(Me.CustomerDescription))
        Sql = Sql.Replace("@station", Me.Station)
        Sql = Sql.Replace("@tolerance_id", Me.AccuracyClass)
        Sql = Sql.Replace("@dual_remote", Me.Database.Escape(Me.DualRemote))
        Sql = Sql.Replace("@inactive", Me.Database.Escape(Me.Inactive))
        Sql = Sql.Replace("@part_no", Me.Database.Escape(Me.PartNo))
        Sql = Sql.Replace("@cal_template_id", Me.CalTemplateId)
        Sql = Sql.Replace("@procedure_id", Me.CalProcedure)
        Sql = Sql.Replace("@tolerance", Me.Database.Escape(Me.Tolerenace))
        Sql = Sql.Replace("@cert_template_id", Me.CertTemplateId)
        If Me.Id = 0 Then
            Me.Database.InsertAndReturnId(Sql)
        Else
            Me.Database.Execute(Sql)
        End If
        If Me.Database.LastQuery.Successful Then
            ' If new, set id
            If Me.Id = 0 Then
                Me._Id = Me.Database.LastQuery.InsertId
            End If
            RaiseEvent Saved(Me)
            ' If it saved successfully check if we need to log a location/status change
            If Me.OldShipTo <> Me.ShipTo Or Me.OldStation <> Me.Station Or Me.OldStatus <> Me.Status Then
                ' Log change
                Sql = "INSERT INTO DEPRECSTATUS (dstat_dep_id, dstat_old_stat, dstat_new_stat, dstat_old_loc, dstat_new_loc,"
                Sql &= " station_id, dstat_chngt, dstat_user)"
                Sql &= " VALUES ("
                Sql &= Me.Database.Escape(Me.Id) & ", "
                Sql &= Me.Database.Escape(Me.OldStatus) & ", "
                Sql &= Me.Database.Escape(Me.Status) & ", "
                Sql &= Me.Database.Escape(Me.OldShipTo) & ", "
                Sql &= Me.Database.Escape(Me.ShipTo) & ", "
                Sql &= Me.Database.Escape(Me.Station) & ", "
                Sql &= " " & Me.Database.Timestamp & ", "
                Sql &= Me.Database.Escape(Me.LastUpdatedBy)
                Sql &= ")"
                Me.Database.Execute(Sql)
                If Not Me.Database.LastQuery.Successful Then
                    MsgBox(Me.Database.LastQuery.ErrorMsg)
                End If
            End If
            ' Reopen
            Me.Open(Me.Id)
        Else
            Throw New Exception("Database error: " & Me.Database.LastQuery.ErrorMsg)
        End If

    End Sub

    Private Function CheckForDuplicate() As Boolean
        Dim Settings As New cSettings(Me.Database)
        If Settings.GetValue("Check for Duplicate Serials", 1) > 0 Then
            If Me.Serial.ToUpper = Settings.GetValue("Equipment No Serial Phrase", "N/A") Then
                Return False
            Else
                Dim sql As String
                Dim count As Integer
                sql = "SELECT COUNT(dep_id) FROM DEPREC WHERE dep_manuf=@make AND dep_mod=@model AND dep_ser=@serial"
                sql = sql.Replace("@make", Me.Database.Escape(Me.Manufacturer))
                sql = sql.Replace("@model", Me.Database.Escape(Me.Model))
                sql = sql.Replace("@serial", Me.Database.Escape(Me.Serial))
                count = Me.Database.GetOne(sql)
                If Me.Id > 0 And count > 1 Then
                    Return True
                ElseIf Me.Id = 0 And count > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
        Else
            Return False
        End If
    End Function

    Private Function IsComplete() As Boolean
        If Me.Category = 0 Then
            Me._ErrorMsg = "Please specify a category."
            Return False
        End If
        If Me.ShipTo = Nothing Then
            Me._ErrorMsg = "Please specify a ship to location."
            Return False
        End If
        If Me.Status = "" Then
            Me._ErrorMsg = "Please specify a status."
            Return False
        End If
        Return True
    End Function

    Private Function IsNull(ByVal Value As Object, ByVal ReturnVal As Object) As Object
        If Value Is DBNull.Value Then
            Return ReturnVal
        Else
            Return Value
        End If
    End Function

    Private Function IsNothing(ByVal Value As Object, ByVal ReturnVal As Object) As Object
        If Value = Nothing Then
            Return ReturnVal
        Else
            Return Value
        End If
    End Function

    Public Function AttachedEquipment() As DataTable
        Dim sql As String
        sql = "SELECT dep_id, dep_type, dep_manuf, dep_mod, dep_ser, "
        sql &= " (dep_manuf + ' ' + dep_mod + ' (' + dep_ser + ')') AS name,"
        sql &= " dep_loc, customer_description, dep_assno,"
        sql &= " dep_cap, dep_CountBy, dep_opt, dep_indicator"
        sql &= " FROM DEPREC"
        sql &= " WHERE (dep_indicator = " & Me.Id
        If Me.AttachedTo > 0 Then
            sql &= " OR dep_id = " & Me.AttachedTo
            sql &= " OR dep_id IN (SELECT dep_id FROM DEPREC WHERE dep_indicator=" & Me.AttachedTo & ")"
        End If
        sql &= ") AND dep_id <> " & Me.Id
        Dim dt As DataTable = Me.Database.GetAll(sql, CommandType.Text)
        If Not Me.Database.LastQuery.Successful Then
            MsgBox(Me.Database.LastQuery.ErrorMsg)
        End If
        Return dt
    End Function



End Class

