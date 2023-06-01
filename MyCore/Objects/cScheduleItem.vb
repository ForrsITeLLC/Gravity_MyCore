Public Class cScheduleItem

    Dim Database As MyCore.Data.EasySql

    Dim _Id As Integer = 0
    Public DateStart As DateTime = Nothing
    Public Duration As Double = 1
    Public CategoryId As Integer = 1
    Public ReferenceId As Integer = 0
    Public Subject As String = ""
    Public Description As String = ""
    Public Owner As String = ""
    Public DateCreated As DateTime = Now
    Public DateLastUpdated As DateTime = Now
    Public CreatedBy As String = ""
    Public LastUpdatedBy As String = ""
    Public Deleted As Boolean = False

    ' Schedule Items only
    Public AllDay As Boolean = False
    Public ReminderInterval As Integer = Nothing
    Public ReminderIntervalUnits As Integer = Nothing
    Public ReminderText As String = Nothing
    Public ReminderSnoozed As Boolean = Nothing
    Public ReminderSnoozeInterval As Integer = Nothing
    Public ReminderSnoozeIntervalUnits As Boolean = Nothing
    Public ReminderSnoozeTime As DateTime = Nothing

    Public IsTimesheet As Boolean = True

    Public Property DurationInMinutes() As Integer
        Get
            Return Me.Duration * 60
        End Get
        Set(ByVal Value As Integer)
            Me.Duration = Value / 60
        End Set
    End Property

    Public Property DurationInHours() As Double
        Get
            Return Me.RoundDuration(Me.DurationInMinutes)
        End Get
        Set(ByVal value As Double)
            Me.Duration = value
        End Set
    End Property

    Public Property Id() As Integer
        Get
            Return Me._Id
        End Get
        Set(ByVal value As Integer)
            Me.Open(value)
        End Set
    End Property

    Public Sub New(ByVal db As MyCore.Data.EasySql, Optional ByVal Timesheet As Boolean = True)
        Me.Database = db
        Me.IsTimesheet = Timesheet
    End Sub

    Public Sub Open(ByVal Id As Integer)
        If Me.IsTimesheet Then
            Dim Sql As String = "SELECT id, category_id, reference_id, subject, description,"
            Sql &= " date_start, duration, owner, date_created, date_last_updated, created_by,"
            Sql &= " last_updated_by, deleted"
            Sql &= " FROM timesheet"
            Sql &= " WHERE id=" & Id
            Dim Row As DataRow
            Try
                Row = Me.Database.GetRow(Sql)
            Catch ex As Exception
                MsgBox(ex.ToString)
                Exit Sub
            End Try
            Me._Id = Id
            Me.CategoryId = Row.Item("category_id")
            Me.ReferenceId = Row.Item("reference_id")
            Me.Subject = Row.Item("subject")
            Me.Description = Row.Item("description")
            Me.DateStart = Row.Item("date_start")
            Me.Duration = Row.Item("duration")
            Me.Owner = Row.Item("owner")
            Me.DateCreated = Row.Item("date_created")
            Me.DateLastUpdated = Row.Item("date_last_updated")
            Me.CreatedBy = Row.Item("created_by")
            Me.LastUpdatedBy = Row.Item("last_updated_by")
            Me.Deleted = Row.Item("deleted")
        Else
            Dim Sql As String = "SELECT * FROM schedule WHERE id=" & Id
            Dim Row As DataRow = Me.Database.GetRow(Sql)
            Me._Id = Id
            Me.CategoryId = Row.Item("category_id")
            Me.ReferenceId = Row.Item("reference_id")
            Me.Subject = Row.Item("subject")
            Me.Description = Row.Item("description")
            Me.DateStart = Row.Item("date_start")
            Me.Duration = Row.Item("duration")
            Me.Owner = Row.Item("owner")
            Me.DateCreated = Row.Item("date_created")
            Me.DateLastUpdated = Row.Item("date_last_updated")
            Me.CreatedBy = Row.Item("created_by")
            Me.LastUpdatedBy = Row.Item("last_updated_by")
            Me.Deleted = Row.Item("deleted")
            Me.AllDay = Row.Item("allday")
            Me.Deleted = Row.Item("deleted")
            ' Yes, below is intentionally all or nothing... if reminder interval is null all should be
            Me.ReminderInterval = IIf(Row.Item("reminder_interval") Is DBNull.Value, Nothing, Row.Item("reminder_interval"))
            Me.ReminderIntervalUnits = IIf(Row.Item("reminder_interval") Is DBNull.Value, Nothing, Row.Item("reminder_interval_units"))
            If Row.Item("reminder_interval") Is DBNull.Value Then
                Me.ReminderSnoozed = Nothing
                Me.ReminderSnoozeInterval = Nothing
                Me.ReminderSnoozeIntervalUnits = Nothing
                Me.ReminderSnoozeTime = Nothing
                Me.ReminderText = Nothing
            Else
                Me.ReminderSnoozed = IIf(Row.Item("reminder_snoozed") Is DBNull.Value, False, Row.Item("reminder_snoozed"))
                If Me.ReminderSnoozed Then
                    Me.ReminderSnoozeInterval = IIf(Row.Item("reminder_snooze_interval") Is DBNull.Value, 15, Row.Item("reminder_snooze_interval"))
                    Me.ReminderSnoozeIntervalUnits = IIf(Row.Item("reminder_snooze_interval_units") Is DBNull.Value, 0, Row.Item("reminder_snooze_interval_units"))
                    Me.ReminderSnoozeTime = IIf(Row.Item("reminder_snooze_time") Is DBNull.Value, Now, Row.Item("reminder_snooze_time"))
                    Me.ReminderText = IIf(Row.Item("reminder_text") Is DBNull.Value, "", Row.Item("reminder_text"))
                Else
                    Me.ReminderSnoozeInterval = Nothing
                    Me.ReminderSnoozeIntervalUnits = Nothing
                    Me.ReminderSnoozeTime = Nothing
                    Me.ReminderText = Nothing
                End If
            End If

        End If
    End Sub

    Public Sub Save()
        If Me.IsTimesheet Then
            If Me._Id > 0 Then
                Dim Sql As String = "UPDATE timesheet SET "
                Sql &= " category_id=" & Me.CategoryId
                Sql &= ", reference_id=" & Me.ReferenceId
                Sql &= ", subject=" & Me.Database.Escape(Me.Subject)
                Sql &= ", description=" & Me.Database.Escape(Me.Description)
                Sql &= ", date_start=" & Me.Database.Escape(Me.DateStart)
                Sql &= ", duration=" & Me.Database.Escape(Me.DurationInHours)
                Sql &= ", owner=" & Me.Database.Escape(Me.Owner)
                Sql &= ", date_last_updated=" & Me.Database.Escape(Now)
                Sql &= ", last_updated_by=" & Me.Database.Escape(Me.LastUpdatedBy)
                Sql &= ", deleted=" & Me.Database.Escape(IIf(Me.Deleted, 1, 0))
                Sql &= " WHERE id=" & Me._Id
                Me.Database.Execute(Sql)
            Else
                Dim Sql As String = "INSERT INTO timesheet ("
                Sql &= "category_id, reference_id, subject, description, date_start, duration, owner, date_last_updated, last_updated_by,"
                Sql &= "date_created, created_by"
                Sql &= ") VALUES (%category_id, %reference_id, %subject, %description, %date_start, %duration, %owner, %date_last_updated, %last_updated_by,"
                Sql &= "%date_last_updated, %last_updated_by)"
                Sql = Sql.Replace("%category_id", Me.CategoryId)
                Sql = Sql.Replace("%reference_id", Me.ReferenceId)
                Sql = Sql.Replace("%subject", Me.Database.Escape(Me.Subject))
                Sql = Sql.Replace("%description", Me.Database.Escape(Me.Description))
                Sql = Sql.Replace("%date_start", Me.Database.Escape(Me.DateStart))
                Sql = Sql.Replace("%duration", Me.Database.Escape(Me.DurationInHours))
                Sql = Sql.Replace("%owner", Me.Database.Escape(Me.Owner))
                Sql = Sql.Replace("%date_last_updated", Me.Database.Escape(Now))
                Sql = Sql.Replace("%last_updated_by", Me.Database.Escape(Me.LastUpdatedBy))
                Me.Database.InsertAndReturnId(Sql)
                If Me.Database.LastQuery.Successful Then
                    Me._Id = Me.Database.LastQuery.InsertId
                Else
                    Throw New Exception("Timesheet entry not saved. " & Me.Database.LastQuery.ErrorMsg)
                    Exit Sub
                End If
            End If
        Else
            If Me._Id > 0 Then
                Dim Sql As String = "UPDATE schedule SET "
                Sql &= " category_id=" & Me.CategoryId
                Sql &= ", reference_id=" & Me.ReferenceId
                Sql &= ", subject=" & Me.Database.Escape(Me.Subject)
                Sql &= ", description=" & Me.Database.Escape(Me.Description)
                Sql &= ", date_start=" & Me.Database.Escape(Me.DateStart)
                Sql &= ", duration=" & Me.Database.Escape(Me.DurationInHours)
                Sql &= ", owner=" & Me.Database.Escape(Me.Owner)
                Sql &= ", date_last_updated=" & Me.Database.Escape(Now)
                Sql &= ", last_updated_by=" & Me.Database.Escape(Me.LastUpdatedBy)
                Sql &= ", deleted=" & Me.Database.Escape(IIf(Me.Deleted, 1, 0))
                Sql &= ", allday=" & Me.Database.Escape(Me.AllDay)
                If Me.ReminderInterval <> Nothing Then
                    Sql &= ", reminder_interval=" & Me.Database.Escape(Me.ReminderInterval)
                    Sql &= ", reminder_interval_units=" & Me.Database.Escape(Me.ReminderIntervalUnits)
                    Sql &= ", reminder_text=" & Me.Database.Escape(Me.ReminderText)
                    If Me.ReminderSnoozed Then
                        Sql &= ", reminder_snoozed=" & Me.Database.Escape(Me.ReminderSnoozed)
                        Sql &= ", reminder_snooze_interval=" & Me.Database.Escape(Me.ReminderSnoozeInterval)
                        Sql &= ", reminder_snooze_interval_units=" & Me.Database.Escape(Me.ReminderSnoozeIntervalUnits)
                        Sql &= ", reminder_snooze_time=" & Me.Database.Escape(Me.ReminderSnoozeTime)
                    Else
                        Sql &= ", reminder_snoozed=" & Me.Database.Escape(Me.ReminderSnoozed)
                        Sql &= ", reminder_snooze_interval=" & Me.Database.Escape(Nothing)
                        Sql &= ", reminder_snooze_interval_units=" & Me.Database.Escape(Nothing)
                        Sql &= ", reminder_snooze_time=" & Me.Database.Escape(Nothing)
                    End If
                Else
                    Sql &= ", reminder_interval=" & Me.Database.Escape(Nothing)
                    Sql &= ", reminder_interval_units=" & Me.Database.Escape(Nothing)
                    Sql &= ", reminder_text=" & Me.Database.Escape(Nothing)
                    Sql &= ", reminder_snoozed=" & Me.Database.Escape(Nothing)
                    Sql &= ", reminder_snooze_interval=" & Me.Database.Escape(Nothing)
                    Sql &= ", reminder_snooze_interval_units=" & Me.Database.Escape(Nothing)
                    Sql &= ", reminder_snooze_time=" & Me.Database.Escape(Nothing)
                End If
                Sql &= " WHERE id=" & Me._Id
                Me.Database.Execute(Sql)
                If Not Me.Database.LastQuery.Successful Then
                    MessageBox.Show(Me.Database.LastQuery.ErrorMsg & " - " & Sql)
                End If
            Else
                Dim Sql As String = "INSERT INTO schedule ("
                Sql &= "category_id, reference_id, subject, description, date_start, duration, owner, date_last_updated, last_updated_by,"
                Sql &= "date_created, created_by,"
                Sql &= "deleted, allday, reminder_interval, reminder_interval_units, reminder_text, reminder_snoozed, reminder_snooze_interval,"
                Sql &= "reminder_snooze_interval_units, reminder_snooze_time"
                Sql &= ") VALUES (%category_id%, %reference_id%, %subject%, %description%, %date_start%, %duration%, %owner%, %date_last_updated%, %last_updated_by%,"
                Sql &= "%date_last_updated%, %last_updated_by%,"
                Sql &= "%deleted%, %allday%, %reminder_interval%, %reminder_interval_units%, %reminder_text%, %reminder_snoozed%, %reminder_snoozed_interval%,"
                Sql &= "%reminder_snoozed_interval_units%, %reminder_snoozed_time%)"
                Sql = Sql.Replace("%category_id%", Me.CategoryId)
                Sql = Sql.Replace("%reference_id%", Me.ReferenceId)
                Sql = Sql.Replace("%subject%", Me.Database.Escape(Me.Subject))
                Sql = Sql.Replace("%description%", Me.Database.Escape(Me.Description))
                Sql = Sql.Replace("%date_start%", Me.Database.Escape(Me.DateStart))
                Sql = Sql.Replace("%duration%", Me.Database.Escape(Me.DurationInHours))
                Sql = Sql.Replace("%owner%", Me.Database.Escape(Me.Owner))
                Sql = Sql.Replace("%date_last_updated%", Me.Database.Escape(Now))
                Sql = Sql.Replace("%last_updated_by%", Me.Database.Escape(Me.LastUpdatedBy))
                Sql = Sql.Replace("%deleted%", Me.Database.Escape(Me.Deleted))
                Sql = Sql.Replace("%allday%", Me.Database.Escape(Me.AllDay))
                If Me.ReminderInterval <> Nothing Then
                    Sql = Sql.Replace("%reminder_interval%", Me.Database.Escape(Me.ReminderInterval))
                    Sql = Sql.Replace("%reminder_interval_units%", Me.Database.Escape(Me.ReminderIntervalUnits))
                    Sql = Sql.Replace("%reminder_text%", Me.Database.Escape(Me.ReminderText))
                    Sql = Sql.Replace("%reminder_snoozed%", Me.Database.Escape(Me.ReminderSnoozed))
                    Sql = Sql.Replace("%reminder_snoozed_interval%", Me.Database.Escape(Me.ReminderSnoozeInterval))
                    Sql = Sql.Replace("%reminder_snoozed_interval_units%", Me.Database.Escape(Me.ReminderSnoozeIntervalUnits))
                    Sql = Sql.Replace("%reminder_snoozed_time%", Me.Database.Escape(Me.ReminderSnoozeTime))
                Else
                    Sql = Sql.Replace("%reminder_interval%", Me.Database.Escape(Nothing))
                    Sql = Sql.Replace("%reminder_interval_units%", Me.Database.Escape(Nothing))
                    Sql = Sql.Replace("%reminder_text%", Me.Database.Escape(Nothing))
                    Sql = Sql.Replace("%reminder_snoozed%", Me.Database.Escape(Nothing))
                    Sql = Sql.Replace("%reminder_snoozed_interval%", Me.Database.Escape(Nothing))
                    Sql = Sql.Replace("%reminder_snoozed_interval_units%", Me.Database.Escape(Nothing))
                    Sql = Sql.Replace("%reminder_snoozed_time%", Me.Database.Escape(Nothing))
                End If
                Me.Database.InsertAndReturnId(Sql)
                If Me.Database.LastQuery.Successful Then
                    Me._Id = Me.Database.LastQuery.InsertId
                Else
                    Throw New Exception("Schedule entry not saved. " & Me.Database.LastQuery.ErrorMsg)
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Public Function ToAppointment() As Infragistics.Win.UltraWinSchedule.Appointment
        Dim ap As New Infragistics.Win.UltraWinSchedule.Appointment(Me.DateStart, Me.DateStart.AddHours(Me.Duration))
        ap.StartDateTime = Me.DateStart
        ap.EndDateTime = Me.DateStart.AddHours(Me.DurationInHours)
        ap.Subject = Me.Subject
        ap.Description = Me.Description
        ap.Tag = Me._Id
        ap.Location = IIf(Me.ReferenceId > 0, Me.ReferenceId, "")
        ap.OwnerKey = Me.Owner
        ap.Reminder.Enabled = False
        ap.OwnerKey = Me.Owner
        ap.DataKey = Me.CategoryId
        ap.Appearance.Key = Me.CategoryId
        If Me.IsTimesheet Then
            ap.AllDayEvent = False
        Else
            ap.AllDayEvent = Me.AllDay
            If Me.ReminderInterval <> Nothing Then
                ap.Reminder.DialogText = Me.ReminderText
                ap.Reminder.DisplayInterval = Me.ReminderInterval
                ap.Reminder.DisplayIntervalUnits = Me.ReminderIntervalUnits
                ap.Reminder.Snoozed = Me.ReminderSnoozed
                ap.Reminder.SnoozeInterval = Me.ReminderSnoozeInterval
                ap.Reminder.SnoozeIntervalUnits = Me.ReminderSnoozeIntervalUnits
                ap.Reminder.SnoozeTime = Me.ReminderSnoozeTime
            End If
        End If
        Return ap
    End Function


    Public Sub FromAppointment(ByVal ap As Infragistics.Win.UltraWinSchedule.Appointment, Optional ByVal SetId As Boolean = False)
        If SetId Then
            Me._Id = ap.Tag
        End If
        If Me.IsTimesheet Then
            Try
                Me.CategoryId = ap.DataKey
            Catch
                Me.CategoryId = 1
            End Try
            Me.DateStart = ap.StartDateTime
            Me.DurationInMinutes = Me.GetDurationMinutes(ap.StartDateTime, ap.EndDateTime)
            Me.Description = ap.Description
            Me.Owner = ap.OwnerKey
            Dim RefId As Integer = 0
            If ap.Location IsNot Nothing Then
                Try
                    RefId = IIf(Microsoft.VisualBasic.IsNumeric(ap.Location), ap.Location, 0)
                Catch
                    RefId = 0
                End Try
            End If
            Me.ReferenceId = RefId
            Me.Subject = ap.Subject
        Else
            Try
                Me.CategoryId = ap.DataKey
            Catch
                Me.CategoryId = 1
            End Try
            Me.DateStart = ap.StartDateTime
            Me.DurationInMinutes = Me.GetDurationMinutes(ap.StartDateTime, ap.EndDateTime)
            Me.Description = ap.Description
            Me.Owner = ap.OwnerKey
            Dim RefId As Integer = 0
            If ap.Location IsNot Nothing Then
                Try
                    RefId = IIf(Microsoft.VisualBasic.IsNumeric(ap.Location), ap.Location, 0)
                Catch
                    RefId = 0
                End Try
            End If
            Me.ReferenceId = RefId
            Me.Subject = ap.Subject
            Me.AllDay = ap.AllDayEvent
            Me.ReminderInterval = ap.Reminder.DisplayInterval
            Me.ReminderIntervalUnits = ap.Reminder.DisplayIntervalUnits
            Me.ReminderSnoozed = ap.Reminder.Snoozed
            Me.ReminderSnoozeInterval = ap.Reminder.SnoozeInterval
            Me.ReminderSnoozeIntervalUnits = ap.Reminder.SnoozeIntervalUnits
            Me.ReminderSnoozeTime = ap.Reminder.SnoozeTime
            Me.ReminderText = ap.Reminder.DialogText
        End If
    End Sub

    Public Function GetDurationMinutes(ByVal Time1 As DateTime, ByVal Time2 As DateTime) As Integer
        Dim Min As Integer = Math.Abs(Microsoft.VisualBasic.DateDiff(DateInterval.Minute, Time1, Time2))
        Return Min
    End Function

    Public Function RoundDuration(ByVal DurationInMinutes As Integer) As Double
        ' Make sure it's positive 
        DurationInMinutes = Math.Abs(DurationInMinutes)
        ' Round down
        DurationInMinutes = Math.Floor(DurationInMinutes)
        ' Get hours and minutes
        Dim Hours As Integer = 0
        Dim Minutes As Integer = 0
        If DurationInMinutes Mod 60 = 0 Then
            Hours = DurationInMinutes / 60
        Else
            Hours = Math.Floor(DurationInMinutes / 60)
            Minutes = Math.Floor(DurationInMinutes Mod 60)
        End If
        ' Round up to next 15 minutes
        If Minutes > 45 Then
            Hours += 1
            Minutes = 0
        ElseIf Minutes > 30 Then
            Minutes = 45
        ElseIf Minutes > 15 Then
            Minutes = 30
        ElseIf Minutes > 0 Then
            Minutes = 15
        Else
            Minutes = 0
        End If
        ' Enforce minimum of 15 minutes... no 0 duration items
        If Minutes = 0 And Hours = 0 Then
            Minutes = 15
        End If
        ' Return value
        Return Hours + (Minutes / 60)
    End Function



End Class
