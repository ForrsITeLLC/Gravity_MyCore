Public Class cTimesheetCalc

    Dim Database As MyCore.Data.EasySql

    Dim s As DateTime
    Dim e As DateTime
    Dim _StandardHours As Double = 0
    Dim _OTHours As Double = 0
    Dim _HolidayHours As Double = 0

    Public WorkDayStart As DateTime
    Public WorkDayEnd As DateTime
    Public WorkDays As DayOfWeek()
    Public HolidayDays As DayOfWeek()
    Public HolidayDates As Date()

    Public ReadOnly Property StandardHours() As Double
        Get
            Return Me._StandardHours
        End Get
    End Property

    Public ReadOnly Property OvertimeHours() As Double
        Get
            Return Me._OTHours
        End Get
    End Property

    Public ReadOnly Property HolidayHours() As Double
        Get
            Return Me._HolidayHours
        End Get
    End Property

    Public ReadOnly Property TotalHours() As Double
        Get
            Return Me._StandardHours + Me._HolidayHours + Me._OTHours
        End Get
    End Property

    Public ReadOnly Property TotalManHours() As Double
        Get
            Return Me._StandardHours + Me._HolidayHours * 2 + Me._OTHours * 1.5
        End Get
    End Property

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
        Dim Holidays As DataTable = Me.Database.GetAll("SELECT id, calendar_date, description FROM holiday ORDER BY calendar_date ASC")
        Dim Days As String = Me.Database.GetOne("SELECT value FROM settings WHERE property='Work Days'")
        Dim HolidayBilling As String = Me.Database.GetOne("SELECT value FROM settings WHERE property='Holiday Billing'")
        Dim DayStart As String = Me.Database.GetOne("SELECT value FROM settings WHERE property='Work Day Start'")
        Dim DayEnd As String = Me.Database.GetOne("SELECT value FROM settings WHERE property='Work Day End'")
        Dim arrDays As String() = Days.Split(",")
        Dim arrHolidayBilling As String() = HolidayBilling.Split(",")
        Dim arrDayStart As String() = DayStart.Split(":")
        Dim arrDayEnd As String() = DayEnd.Split(":")
        ' Holidays
        ReDim Me.HolidayDates(Holidays.Rows.Count - 1)
        For i As Integer = 0 To Holidays.Rows.Count - 1
            Me.HolidayDates(i) = Holidays.Rows(i).Item("calendar_date")
        Next
        ' Work Hours
        Me.WorkDayStart = New DateTime(Today.Year, Today.Month, Today.Day, arrDayStart(0), arrDayStart(1), 0)
        Me.WorkDayEnd = New DateTime(Today.Year, Today.Month, Today.Day, arrDayEnd(0), arrDayEnd(1), 0)
        ' Work Days
        ReDim Me.WorkDays(arrDays.Length - 1)
        For i As Integer = 0 To arrDays.Length - 1
            Select Case arrDays(i)
                Case "Mon"
                    Me.WorkDays(i) = DayOfWeek.Monday
                Case "Tue"
                    Me.WorkDays(i) = DayOfWeek.Tuesday
                Case "Wed"
                    Me.WorkDays(i) = DayOfWeek.Wednesday
                Case "Thu"
                    Me.WorkDays(i) = DayOfWeek.Thursday
                Case "Fri"
                    Me.WorkDays(i) = DayOfWeek.Friday
                Case "Sat"
                    Me.WorkDays(i) = DayOfWeek.Saturday
                Case "Sun"
                    Me.WorkDays(i) = DayOfWeek.Sunday
            End Select
        Next
        ' Holiday billing days
        ReDim Me.HolidayDays(arrHolidayBilling.Length - 1)
        For i As Integer = 0 To arrHolidayBilling.Length - 1
            Select Case arrHolidayBilling(i)
                Case "Sat"
                    Me.HolidayDays(i) = DayOfWeek.Saturday
                Case "Sun"
                    Me.HolidayDays(i) = DayOfWeek.Sunday
            End Select
        Next
    End Sub

    Public Sub SetTimesheetEntry(ByVal InTime As DateTime, ByVal OutTime As DateTime)
        Me.s = InTime
        Me.e = OutTime
        Me.SplitHours()
    End Sub

    Private Function isSingleDay() As Boolean
        If s.Year = e.Year And s.Month = e.Month And s.Day = e.Day Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function GetDaysInSpan(ByVal Start As Date, ByVal Finish As Date) As Collection
        Dim Current As New Date(Start.Year, Start.Month, Start.Day, 0, 0, 0)
        Dim Dates As New Collection
        Dim i As Integer = 0
        While Current.Ticks <= Finish.Ticks
            Dates.Add(New Date(Current.Year, Current.Month, Current.Day, 0, 0, 0))
            Current = Current.AddDays(1)
            i += 1
        End While
        Return Dates
    End Function

    Private Function IsWeekend(ByVal Day As Date) As Boolean
        For i As Integer = 0 To Me.WorkDays.Length - 1
            If Day.DayOfWeek.ToString = Me.WorkDays(i).ToString Then
                Return False
            End If
        Next
        Return True
    End Function

    Private Function IsHoliday(ByVal Day As Date) As Boolean
        For Each Holiday As Date In Me.HolidayDates
            If Holiday = Day Then
                Return True
            End If
        Next
        For Each WorkDay As DayOfWeek In Me.WorkDays
            If Day.DayOfWeek = WorkDay Then
                Return False
            End If
        Next
        Return True
    End Function

    Public Function GetPricePerHour(ByVal TotalCharge As Double) As Double
        Return TotalCharge / Me.TotalManHours
    End Function

    Public Sub SplitHours()

        Dim TotalMinutes As Integer = Microsoft.VisualBasic.DateDiff(DateInterval.Minute, Me.s, Me.e)
        Dim StandardMinutes As Integer = 0
        Dim OTMinutes As Integer = 0
        Dim HolidayMinutes As Integer = 0
        Dim ReturnMinutes(2) As Integer

        If Me.isSingleDay Then

            If Not Me.IsHoliday(Me.s) Then

                ' Beginning of Day OT
                If Me.s.Hour < Me.WorkDayStart.Hour Or (Me.s.Hour = Me.WorkDayStart.Hour And Me.s.Minute < Me.WorkDayStart.Minute) Then
                    OTMinutes += (Me.WorkDayStart.Hour * 60 + Me.WorkDayStart.Minute) - (Me.s.Hour * 60 + Me.s.Minute)
                End If

                ' End of day OT
                If Me.e.Hour > Me.WorkDayEnd.Hour Or (Me.e.Hour = Me.WorkDayEnd.Hour And Me.e.Minute > Me.WorkDayEnd.Minute) Then
                    OTMinutes += (Me.e.Hour * 60 + Me.e.Minute) - (Me.WorkDayEnd.Hour * 60 + Me.WorkDayEnd.Minute)
                End If

            Else

                HolidayMinutes += Microsoft.VisualBasic.DateDiff(DateInterval.Minute, Me.s, Me.e)

            End If

            StandardMinutes = TotalMinutes - OTMinutes - HolidayMinutes

        Else

            Dim Days As Collection = Me.GetDaysInSpan(Me.s, Me.e)

            For i As Integer = 1 To Days.Count

                Dim Current As Date = Days.Item(i)

                If Not Me.IsHoliday(Current) Then

                    If Current.Date = Me.s.Date Then

                        If Me.s.Hour < Me.WorkDayStart.Hour Or (Me.s.Hour = Me.WorkDayStart.Hour And Me.s.Minute < Me.WorkDayStart.Minute) Then
                            OTMinutes += (Me.WorkDayStart.Hour * 60 + Me.WorkDayStart.Minute) - (Me.s.Hour * 60 + Me.s.Minute)
                        End If

                        If Me.s.Hour < Me.WorkDayEnd.Hour Or (Me.s.Hour = Me.WorkDayEnd.Hour And Me.s.Minute < Me.WorkDayEnd.Minute) Then
                            OTMinutes += 24 * 60 - (Me.WorkDayEnd.Hour * 60 + Me.WorkDayEnd.Minute)
                        Else
                            OTMinutes += 24 * 60 - (Me.s.Hour * 60 + Me.s.Minute)
                        End If

                    ElseIf Current.Date = Me.e.Date Then

                        If Me.e.Hour < Me.WorkDayStart.Hour Or (Me.e.Hour = Me.WorkDayStart.Hour And Me.e.Minute < Me.WorkDayStart.Minute) Then
                            OTMinutes += Me.e.Hour * 60 + Me.e.Minute
                        Else
                            OTMinutes += Me.WorkDayStart.Hour * 60 + Me.WorkDayStart.Minute
                        End If

                        If Me.e.Hour > Me.WorkDayEnd.Hour Or (Me.e.Hour = Me.WorkDayEnd.Hour And Me.e.Minute > Me.WorkDayEnd.Minute) Then
                            OTMinutes += (Me.e.Hour * 60 + Me.e.Minute) - (Me.WorkDayEnd.Hour * 60 + Me.WorkDayEnd.Minute)
                        End If

                    Else

                        OTMinutes += Me.WorkDayStart.Hour * 60 + Me.WorkDayStart.Minute
                        OTMinutes += 24 * 60 - (Me.WorkDayEnd.Hour * 60 + Me.WorkDayEnd.Minute)

                    End If

                Else

                    If Current.Date = Me.s.Date Then
                        HolidayMinutes += 24 * 60 - (Me.s.Hour * 60 + Me.s.Minute)
                    ElseIf Current.Date = Me.e.Date Then
                        HolidayMinutes += Me.s.Hour * 60 + Me.s.Minute
                    Else
                        HolidayMinutes += 24 * 60
                    End If

                End If

            Next

            StandardMinutes = TotalMinutes - OTMinutes - HolidayMinutes

        End If

        Me._StandardHours = Math.Round(StandardMinutes / 60, 2)
        Me._OTHours = Math.Round(OTMinutes / 60, 2)
        Me._HolidayHours = Math.Round(HolidayMinutes / 60, 2)

    End Sub





End Class
