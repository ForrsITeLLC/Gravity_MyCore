Public Class cSettings

    Public Database As MyCore.Data.EasySql

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
    End Sub

    Public Function GetValue(ByVal PropName As String, Optional ByVal DefaultVal As String = "") As String
        Dim Value As String = Me.Database.GetOne("SELECT value FROM settings WHERE property=" & Me.Database.Escape(PropName))
        If Me.Database.LastQuery.RowsReturned = 0 Then
            Value = DefaultVal
            Dim Sql As String = "INSERT INTO settings (property, value, date_last_updated) VALUES ("
            Sql &= Me.Database.Escape(PropName) & ", "
            Sql &= Me.Database.Escape(Value) & ", "
            Sql &= Me.Database.Escape(Now) & ")"
            Me.Database.Execute(Sql)
        End If
        Return Value
    End Function

    Public Sub SetValue(ByVal PropName As String, ByVal Value As String)
        Me.Database.Execute("UPDATE settings SET value=" & Me.Database.Escape(Value) & ", date_last_updated=" & Me.Database.Escape(Now) & " WHERE property=" & Me.Database.Escape(PropName))
        If Not Me.Database.LastQuery.Successful Then
            Dim Err As String = Me.Database.LastQuery.ErrorMsg
        End If
    End Sub

End Class
