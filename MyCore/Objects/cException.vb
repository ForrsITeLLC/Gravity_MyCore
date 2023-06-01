Public Class cException
    Inherits Exception

    Public Shadows Message As String = ""
    Public Details As String = ""
    Public Title As String = "Error"
    Public Severity As SeverityRating = SeverityRating.Minor

    Public Enum SeverityRating
        Informational = 0
        Minor = 1
        Serious = 2
        Critical = 3
        Fatal = 4
    End Enum

    Public Sub New(ByVal Severity As cException.SeverityRating, ByVal Message As String, Optional ByVal Details As String = "", Optional ByVal Title As String = "Error")
        Me.Severity = Severity
        Me.Message = Message
        Me.Details = Details
        Me.Title = Title
    End Sub

End Class
