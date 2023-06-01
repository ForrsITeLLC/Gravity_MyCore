Imports MyCore.Data

Public Class cStandard

    Dim Database As MyCore.Data.EasySql

    Dim _Id As Integer = 0

    Public AssetNo As String = ""
    Public Manufacturer As String = ""
    Public Model As String = ""
    Public Notes As String = ""
    Public SerialNo As String = ""
    Public StationId As Integer = 0
    Public DateLastInternalCal As DateTime = Nothing
    Public InternalCalFrequency As Integer = 0
    Public Removed As Boolean = False

    Public ReadOnly Property Id() As Integer
        Get
            Return Me._Id
        End Get
    End Property

    Public Event Reload()
    Public Event Saved(ByVal Standard As cStandard)

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
    End Sub

    Public Sub Open(ByVal AssetNo As String)
        Dim Sql As String = "SELECT * FROM standards_equipment WHERE asset_no='" & AssetNo & "'"
        Dim Row As DataRow = Me.Database.GetRow(Sql)
        If Me.Database.LastQuery.RowsReturned > 0 Then
            Me._Id = Row.Item("id")
            Me.AssetNo = AssetNo
            Me.Manufacturer = Row.Item("manufacturer")
            Me.Model = Row.Item("model")
            Me.SerialNo = Row.Item("serial_no")
            Me.Notes = Row.Item("description")
            Me.Removed = Row.Item("removed")
            Me.InternalCalFrequency = Row.Item("internal_cal_frequency")
            If Row.Item("date_last_internal_cal") IsNot DBNull.Value Then
                Me.DateLastInternalCal = Row.Item("date_last_internal_cal")
            Else
                Me.DateLastInternalCal = Nothing
            End If
            Me.StationId = Row.Item("station_id")
            RaiseEvent Reload()
        Else
            Throw New Exception("Standard does not exist.")
        End If
    End Sub

    Public Sub OpenAsNew()
        Me._Id = 0
        RaiseEvent Reload()
    End Sub

    Public Sub Save()
        If Me.Id = 0 Then
            Dim Sql As String = "INSERT INTO standards_equipment (asset_no, serial_no, manufacturer, model, station_id,"
            Sql &= " date_last_internal_cal, internal_cal_frequency, description, date_last_updated)"
            Sql &= " VALUES (@asset_no, @serial_no, @manufacturer, @model, @station_id, "
            Sql &= " @date_last_internal_cal, @internal_cal_frequency, @description, " & Me.Database.Timestamp & ")"
            Sql = Sql.Replace("@asset_no", Me.Database.Escape(Me.AssetNo))
            Sql = Sql.Replace("@serial_no", Me.Database.Escape(Me.SerialNo))
            Sql = Sql.Replace("@manufacturer", Me.Database.Escape(Me.Manufacturer))
            Sql = Sql.Replace("@model", Me.Database.Escape(Me.Model))
            Sql = Sql.Replace("@station_id", Me.StationId)
            Sql = Sql.Replace("@description", Me.Database.Escape(Me.Notes))
            Sql = Sql.Replace("@internal_cal_frequency", Me.InternalCalFrequency)
            Sql = Sql.Replace("@date_last_internal_cal", Me.Database.Escape(IIf(Me.DateLastInternalCal = Nothing, DBNull.Value, Me.DateLastInternalCal)))
            Sql = Sql.Replace("@removed", Me.Database.Escape(Me.Removed))
            Me.Database.InsertAndReturnId(Sql)
            If Me.Database.LastQuery.Successful Then
                Me._Id = Me.Database.LastQuery.InsertId
            Else
                Throw New Exception("Standard not saved. " & Me.Database.LastQuery.ErrorMsg)
                Exit Sub
            End If
        Else
            Dim Sql As String = "UPDATE standards_equipment"
            Sql &= " SET asset_no=@asset_no, serial_no=@serial_no, manufacturer=@manufacturer, model=@model, station_id=@station_id,"
            Sql &= " date_last_internal_cal=@date_last_internal_cal, internal_cal_frequency=@internal_cal_frequency,"
            Sql &= " description=@description, date_last_updated=" & Me.Database.Timestamp & ", removed=@removed"
            Sql &= " WHERE id=@id"
            Sql = Sql.Replace("@asset_no", Me.Database.Escape(Me.AssetNo))
            Sql = Sql.Replace("@serial_no", Me.Database.Escape(Me.SerialNo))
            Sql = Sql.Replace("@manufacturer", Me.Database.Escape(Me.Manufacturer))
            Sql = Sql.Replace("@model", Me.Database.Escape(Me.Model))
            Sql = Sql.Replace("@station_id", Me.StationId)
            Sql = Sql.Replace("@description", Me.Database.Escape(Me.Notes))
            Sql = Sql.Replace("@internal_cal_frequency", Me.InternalCalFrequency)
            Sql = Sql.Replace("@date_last_internal_cal", Me.Database.Escape(IIf(Me.DateLastInternalCal = Nothing, DBNull.Value, Me.DateLastInternalCal)))
            Sql = Sql.Replace("@removed", Me.Database.Escape(Me.Removed))
            Sql = Sql.Replace("@id", Me.Id)
            Me.Database.Execute(Sql)
            If Not Me.Database.LastQuery.Successful Then
                Throw New Exception("Standard not saved. " & Me.Database.LastQuery.ErrorMsg)
                Exit Sub
            End If
        End If
        RaiseEvent Saved(Me)
        Me.Open(Me.AssetNo)
    End Sub

    Public Sub ChangeAssetNo(ByVal NewAssetNo As String)
        Dim Sql As String = ""
        ' First, Check if NewAssetNo already exists... bail out it does
        Sql = "SELECT COUNT(id) FROM standards_equipment WHERE asset_no='" & NewAssetNo & "' GROUP BY asset_no"
        If Me.Database.GetOne(Sql) > 0 Then
            Throw New Exception("Asset No already exists.")
            Exit Sub
        End If
        ' Change standards used records on work orders
        Sql = "UPDATE standards_used SET asset_no='" & NewAssetNo & "' WHERE asset_no='" & Me.AssetNo & "'"
        Me.Database.Execute(Sql)
        ' Change all standards certifications
        Sql = "UPDATE standards_to_cert SET asset_no='" & NewAssetNo & "' WHERE asset_no='" & Me.AssetNo & "'"
        Me.Database.Execute(Sql)
        If Not Me.Database.LastQuery.Successful Then
            Throw New cException(cException.SeverityRating.Serious, "Error changing tests to new asset no.", Me.Database.LastQuery.ErrorMsg & " " & Me.Database.LastQuery.CommandText)
        End If
        ' Finally, Change standard record
        Sql = "UPDATE standards_equipment SET asset_no='" & NewAssetNo & "' WHERE id=" & Me.Id
        Me.Database.Execute(Sql)
        If Me.Database.LastQuery.Successful Then
            Me.AssetNo = NewAssetNo
        Else
            Throw New Exception(Me.Database.LastQuery.ErrorMsg)
        End If
    End Sub

    Public Function StateCertifications() As DataTable
        Dim Sql As String = "SELECT * FROM standards_certification"
        Sql &= " WHERE test_no IN (SELECT test_no FROM standards_to_cert WHERE asset_no='" & Me.AssetNo & "')"
        Dim dt As DataTable = Me.Database.GetAll(Sql)
        If Me.Database.LastQuery.Successful Then
            Return dt
        Else
            Dim msg As String = Me.Database.LastQuery.ErrorMsg
            Return Nothing
        End If
    End Function

End Class
