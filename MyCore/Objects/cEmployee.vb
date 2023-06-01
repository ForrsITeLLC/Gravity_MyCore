Public Class cEmployee

    Public Id As Integer = Nothing
    Public FirstName As String = ""
    Public LastName As String = ""
    Public UserName As String = ""
    Public Email As String = ""
    Public Permission As Integer = 0
    Public Type As Integer = 1
    Public Deactivated As Boolean = False
    Public PasswordHash As String = ""
    Public Office As String = ""
    Public EmailSig As String = ""
    Public OfficePhone As String = ""
    Public CellPhone As String = ""
    Public Fax As String = ""
    Public IMName As String = ""
    Public VOIPNAme As String = ""

    Dim Database As MyCore.Data.EasySql

    Public Event Reload()
    Public Event Saved(ByVal Employee As cEmployee)


    Public Sub New(ByRef db As MyCore.Data.EasySql)
        Me.Database = db
    End Sub

    Public Sub Open(ByVal Id As Integer)
        Dim Sql As String = "SELECT * FROM employee WHERE id=" & Id
        Dim Row As DataRow = Me.Database.GetRow(Sql)
        Me.Id = Id
        Me.FirstName = Me.IsNull(Row.Item("first_name"))
        Me.LastName = Me.IsNull(Row.Item("last_name"))
        Me.UserName = Me.IsNull(Row.Item("windows_user"))
        Me.Permission = Me.IsNull(Row.Item("perms"), 0)
        Me.Type = Me.IsNull(Row.Item("type"), 1)
        Me.Deactivated = IIf(Row.Item("deactivated"), True, False)
        Me.PasswordHash = Me.IsNull(Row.Item("password_hash"))
        Me.Email = Me.IsNull(Row.Item("email_address"))
        Try
            Me.EmailSig = Me.IsNull(Row.Item("email_signature"))
            Me.CellPhone = Me.IsNull(Row.Item("phone_cell"))
            Me.OfficePhone = Me.IsNull(Row.Item("phone_office"))
            Me.Fax = Me.IsNull(Row.Item("fax"))
            Me.IMName = Me.IsNull(Row.Item("im_name"))
            Me.VOIPNAme = Me.IsNull(Row.Item("voip_name"))
            Me.Office = Me.IsNull(Row.Item("office"))
        Catch ex As Exception
            ' Ignore new fields
        End Try
        RaiseEvent Reload()
    End Sub

    Public Sub OpenByUserName(ByVal User As String)
        Dim Sql As String = "SELECT * FROM employee WHERE windows_user=" & Me.Database.Escape(User)
        Dim Row As DataRow = Me.Database.GetRow(Sql)
        Me.Id = Row.Item("id")
        Me.FirstName = Me.IsNull(Row.Item("first_name"))
        Me.LastName = Me.IsNull(Row.Item("last_name"))
        Me.UserName = Me.IsNull(Row.Item("windows_user"))
        Me.Permission = Me.IsNull(Row.Item("perms"), 0)
        Me.Type = Me.IsNull(Row.Item("type"), 1)
        Me.Deactivated = IIf(Row.Item("deactivated"), True, False)
        Me.PasswordHash = Me.IsNull(Row.Item("password_hash"))
        Me.Email = Me.IsNull(Row.Item("email_address"))
        Try
            Me.EmailSig = Me.IsNull(Row.Item("email_signature"))
            Me.CellPhone = Me.IsNull(Row.Item("phone_cell"))
            Me.OfficePhone = Me.IsNull(Row.Item("phone_office"))
            Me.Fax = Me.IsNull(Row.Item("fax"))
            Me.IMName = Me.IsNull(Row.Item("im_name"))
            Me.VOIPNAme = Me.IsNull(Row.Item("voip_name"))
            Me.Office = Me.IsNull(Row.Item("office"))
        Catch ex As Exception
            ' Ignore new fields
        End Try
        RaiseEvent Reload()
    End Sub

    Public Function Save() As Boolean
        ' Make sure no user name is not already in use
        Dim Sql2 As String = "SELECT COUNT(id) FROM employee WHERE windows_user=" & Me.Database.Escape(Me.UserName)
        If Me.Id <> Nothing Then
            Sql2 &= " AND id <> " & Me.Id & " GROUP BY windows_user"
        End If
        Dim C As Integer = Me.Database.GetOne(Sql2)
        If C = 0 Then
            If Me.Id > 0 Then
                ' Edit
                Dim Sql As String = "UPDATE employee SET"
                Sql &= " windows_user=" & Me.Database.Escape(Me.UserName) & ", "
                Sql &= " first_name=" & Me.Database.Escape(Me.FirstName) & ", "
                Sql &= " last_name=" & Me.Database.Escape(Me.LastName) & ", "
                Sql &= " type=" & Me.Database.Escape(Me.Type) & ", "
                Sql &= " perms=" & Me.Database.Escape(Me.Permission) & ", "
                Sql &= " deactivated=" & IIf(Me.Deactivated, 1, 0) & ", "
                Sql &= " password_hash=" & Me.Database.Escape(Me.PasswordHash) & ", "
                Sql &= " email_signature=" & Me.Database.Escape(Me.EmailSig) & ", "
                Sql &= " phone_cell=" & Me.Database.Escape(Me.CellPhone) & ", "
                Sql &= " phone_office=" & Me.Database.Escape(Me.OfficePhone) & ", "
                Sql &= " fax=" & Me.Database.Escape(Me.Fax) & ", "
                Sql &= " im_name=" & Me.Database.Escape(Me.IMName) & ", "
                Sql &= " voip_name=" & Me.Database.Escape(Me.VOIPNAme) & ", "
                Sql &= " office=" & Me.Database.Escape(Me.Office) & ", "
                Sql &= " email_address=" & Me.Database.Escape(Me.Email)
                Sql &= " WHERE id=" & Me.Id
                Me.Database.Execute(Sql)
            Else
                '  New
                Dim Sql As String = "INSERT INTO employee (windows_user, first_name, last_name,"
                Sql &= " type, perms, deactivated, password_hash, email_address, email_signature,"
                Sql &= " phone_cell, phone_office, fax, voip_name, im_name, office) VALUES ("
                Sql &= Me.Database.Escape(Me.UserName) & ", "
                Sql &= Me.Database.Escape(Me.FirstName) & ", "
                Sql &= Me.Database.Escape(Me.LastName) & ", "
                Sql &= Me.Database.Escape(Me.Type) & ", "
                Sql &= Me.Database.Escape(Me.Permission) & ", "
                Sql &= IIf(Me.Deactivated, 1, 0) & ", "
                Sql &= Me.Database.Escape(Me.PasswordHash) & ", "
                Sql &= Me.Database.Escape(Me.Email) & ", "
                Sql &= Me.Database.Escape(Me.EmailSig) & ", "
                Sql &= Me.Database.Escape(Me.CellPhone) & ", "
                Sql &= Me.Database.Escape(Me.OfficePhone) & ", "
                Sql &= Me.Database.Escape(Me.Fax) & ", "
                Sql &= Me.Database.Escape(Me.VOIPNAme) & ", "
                Sql &= Me.Database.Escape(Me.IMName) & ", "
                Sql &= Me.Database.Escape(Me.Office)
                Sql &= ")"
                Me.Database.InsertAndReturnId(Sql)
                If Me.Database.LastQuery.Successful Then
                    Me.Open(Me.Database.LastQuery.InsertId)
                End If
            End If
            If Me.Database.LastQuery.Successful Then
                RaiseEvent Saved(Me)
            Else
                Throw New Exception(Me.Database.LastQuery.ErrorMsg)
            End If
        Else
            Throw New Exception("Not saved because that user name is already in use.")
        End If
    End Function

    Private Function IsNull(ByVal Item As Object, Optional ByVal DefaultVal As String = "") As Object
        If Item Is DBNull.Value Then
            Return DefaultVal
        Else
            Return Item
        End If
    End Function

End Class
