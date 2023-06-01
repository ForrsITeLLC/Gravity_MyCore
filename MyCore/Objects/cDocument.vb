Public Class cDocument

    Dim _InternalCount As Integer = 0
    Dim Pages As New Hashtable

    Public ReadOnly Property PageCount() As Integer
        Get
            Return Pages.Count
        End Get
    End Property

    Public Sub Add(ByVal Page As String)
        Pages.Add(Me._InternalCount, New cDocumentPage(Page))
        Me._InternalCount += 1
    End Sub

    Public Function Page(ByVal n As Integer) As cDocumentPage
        Dim i As Integer = 0
        For Each p As cDocumentPage In Pages
            If n = i Then
                Return p
            End If
            i += 1
        Next
        Throw New Exception("Page not found.")
    End Function


End Class
