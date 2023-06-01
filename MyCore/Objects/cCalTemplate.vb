Public Class cCalTemplate

    Dim _Id As Integer = 0
    Dim _LastUpdated As DateTime
    Public Name As String = ""
    Public Description As String = ""
    Public AllowOverride As Boolean = True
    Dim Database As MyCore.Data.EasySql

    Public Points As DataTable

    Public Event Reload()
    Public Event CreatedNew()

    Public ReadOnly Property Id() As Integer
        Get
            Return Me._Id
        End Get
    End Property

    Public ReadOnly Property DateLastUpdated() As DateTime
        Get
            Return Me._LastUpdated
        End Get
    End Property

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
    End Sub

    Public Sub Open(ByVal Id As Integer)
        Dim r As DataRow = Me.Database.GetRow("SELECT * FROM calibration_template WHERE id=" & Id)
        If Me.Database.LastQuery.RowsReturned = 1 Then
            Me._Id = Id
            Me._LastUpdated = r.Item("date_last_updated")
            Me.Name = r.Item("name")
            Me.Description = r.Item("description")
            Me.AllowOverride = r.Item("allow_override")
            Me.Points = Me.Database.GetAll("SELECT * FROM calibration_template_point WHERE cal_template_id=" & Id)
        Else
            Throw New Exception("Calibration template not found.")
            Exit Sub
        End If
        RaiseEvent Reload()
    End Sub

    Public Sub OpenAsNew()
        Me.Points = New DataTable
        Me.Points.Columns.Add("id")
        Me.Points.Columns.Add("cal_template_id")
        Me.Points.Columns.Add("test_name")
        Me.Points.Columns.Add("applied_value")
        Me.Points.Columns.Add("applied_units")
        Me.Points.Columns.Add("tolerance_value")
        Me.Points.Columns.Add("date_last_updated")
        RaiseEvent Reload()
    End Sub

    Public Sub Save()
        Dim blnNew As Boolean = False
        If Me._Id = 0 Then
            Dim Sql As String = "INSERT INTO calibration_template (name, description, allow_override, date_last_updated)"
            Sql &= " VALUES (" & Me.Database.Escape(Me.Name) & ", " & Me.Database.Escape(Me.Description)
            Sql &= ", " & Me.Database.Escape(Me.AllowOverride)
            Sql &= ", " & Me.Database.Escape(Now) & ")"
            Me.Database.InsertAndReturnId(Sql)
            If Me.Database.LastQuery.Successful Then
                Me._Id = Me.Database.LastQuery.InsertId
                blnNew = True
            Else
                Throw New Exception(Me.Database.LastQuery.ErrorMsg)
                Exit Sub
            End If
        Else
            Dim Sql As String = "UPDATE calibration_template SET"
            Sql &= " name=" & Me.Database.Escape(Me.Name)
            Sql &= ", description=" & Me.Database.Escape(Me.Description)
            Sql &= ", allow_override=" & Me.Database.Escape(Me.AllowOverride)
            Sql &= ", date_last_updated=" & Me.Database.Escape(Now)
            Sql &= " WHERE id=" & Me.Id
            Me.Database.Execute(Sql)
            If Not Me.Database.LastQuery.Successful Then
                Throw New Exception(Me.Database.LastQuery.ErrorMsg)
                Exit Sub
            End If
        End If
        For Each r As DataRow In Me.Points.Rows
            If r.RowState = DataRowState.Added Then
                Dim Sql As String = "INSERT INTO calibration_template_point (cal_template_id, test_name, applied_value, applied_units, tolerance_value, date_last_updated)"
                Sql &= " VALUES (" & Me.Id & ", " & Me.Database.Escape(r.Item("test_name")) & ", "
                Sql &= Me.Database.Escape(r.Item("applied_value")) & ", " & Me.Database.Escape(r.Item("applied_units")) & ", " & Me.Database.Escape(r.Item("tolerance_value"))
                Sql &= ", " & Me.Database.Escape(Now) & ")"
                Me.Database.Execute(Sql)
            ElseIf r.RowState = DataRowState.Modified Then
                Dim Sql As String = "UPDATE calibration_template_point SET"
                Sql &= " test_name=" & Me.Database.Escape(r.Item("test_name")) & ","
                Sql &= " applied_value=" & Me.Database.Escape(r.Item("applied_value")) & ","
                Sql &= " applied_units=" & Me.Database.Escape(r.Item("applied_units")) & ","
                Sql &= " tolerance_value=" & Me.Database.Escape(r.Item("tolerance_value")) & ","
                Sql &= " date_last_updated=" & Me.Database.Escape(Now)
                Sql &= " WHERE id=" & r.Item("id")
                Me.Database.Execute(Sql)
            End If
        Next
        Me.Open(Me.Id)
        RaiseEvent Reload()
        If blnNew Then
            RaiseEvent CreatedNew()
        End If
    End Sub

    Public Sub DeleteTestPoint(ByVal id As Integer)
        Dim Sql As String = "DELETE FROM calibration_template_point"
        Sql &= " WHERE id=" & id
        Me.Database.Execute(Sql)
    End Sub

    Public Function Units() As DataTable
        Return Me.Database.GetAll("SELECT * FROM units_of_measure ORDER BY sort, name")
    End Function

End Class
