Public Class cCalProcedure

    Dim Database As MyCore.Data.EasySql
    Dim _Id As Integer = 0

    Public Name As String = ""
    Public Content As String = ""

    Public Event Reload()

    Public ReadOnly Property Id() As Integer
        Get
            Return Me._Id
        End Get
    End Property

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
    End Sub

    Public Sub OpenAsNew()
        Me._Id = 0
        Me.Name = ""
        Me.Content = ""
        RaiseEvent Reload()
    End Sub

    Public Sub Open(ByVal Id As Integer)
        Dim Sql As String = "SELECT * FROM cal_procedure WHERE id=" & Me.Database.Escape(Id)
        Dim Row As DataRow = Me.Database.GetRow(Sql)
        If Not Me.Database.LastQuery.Successful Then
            Throw New Exception("Could not open calibration procedure. " & Me.Database.LastQuery.ErrorMsg)
        ElseIf Me.Database.LastQuery.RowsReturned = 0 Then
            Throw New Exception("Did not find calibration procedure with ID# " & Id)
        Else
            Me._Id = Id
            Me.Name = Row.Item("name")
            Me.Content = Row.Item("content")
            RaiseEvent Reload()
        End If
    End Sub

    Public Sub Save()
        Dim Sql As String = ""
        ' Make sure name is not taken
        Dim Matches As DataTable = Me.Database.GetAll("SELECT id FROM cal_procedure WHERE name=" & Me.Database.Escape(Me.Name))
        If Matches.Rows.Count > 1 Then
            Throw New Exception("That name is already taken.  Each calibration procedure must have a unique name.")
        ElseIf Matches.Rows.Count = 1 And Me._Id = 0 Then
            Throw New Exception("That name is already taken.  Each calibration procedure must have a unique name.")
        ElseIf Matches.Rows.Count = 1 Then
            If Matches.Rows(0).Item("id") <> Me._Id Then
                Throw New Exception("That name is already taken.  Each calibration procedure must have a unique name.")
            End If
        End If
        ' Save
        If Me._Id = 0 Then
            ' New
            Sql &= "INSERT INTO cal_procedure (name, content, date_last_updated)"
            Sql &= " VALUES (" & Me.Database.Escape(Me.Name) & ", "
            Sql &= Me.Database.Escape(Me.Content) & ", "
            Sql &= Me.Database.Escape(Now) & ")"
            Me.Database.InsertAndReturnId(Sql)
            RaiseEvent Reload()
        Else
            ' Updated
            Sql &= "UPDATE cal_procedure SET"
            Sql &= " name=" & Me.Database.Escape(Me.Name) & ","
            Sql &= " content=" & Me.Database.Escape(Me.Content) & ","
            Sql &= " date_last_updated=" & Me.Database.Escape(Now)
            Sql &= " WHERE id=" & Me.Id
            Me.Database.Execute(Sql)
        End If
        If Not Me.Database.LastQuery.Successful Then
            Throw New Exception("Could not save calibration procedure. " & Me.Database.LastQuery.ErrorMsg)
        ElseIf Me._Id = 0 Then
            Me._Id = Me.Database.LastQuery.InsertId
            RaiseEvent Reload()
        End If
    End Sub

End Class
