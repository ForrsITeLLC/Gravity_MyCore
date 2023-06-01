Public Class cElement

    Public Id As String = ""
    Public Text As String = ""
    Public Left As Integer = 0
    Public Top As Integer = 0
    Public ZIndex As Integer = 0
    Public Width As Integer = 50
    Public Height As Integer = 50
    Public BorderStyle As BorderStyleType = BorderStyleType.None
    Public BorderWidth As Integer = 0
    Public BorderColor As System.Drawing.Color = Color.Transparent
    Public BackgroundColor As System.Drawing.Color = Color.Transparent
    Friend _ClassName As String

    Public ReadOnly Property ClassName() As String
        Get
            Return Me._ClassName
        End Get
    End Property

    Public Enum BorderStyleType
        None = 0
        Solid = 1
        Ridge = 2
    End Enum

    Public Function StyleString() As String
        Dim Style As String = ""
        Style = "position: absolute; "
        Style &= " left: " & Me.Left & "; "
        Style &= " top: " & Me.Top & "; "
        Style &= " width: " & Me.Width & "; "
        Style &= " height: " & Me.Height & "; "
        If Me.BackgroundColor <> Color.Transparent Then
            Style &= "background-color: rgb(" & Me.BackgroundColor.R & ", " & Me.BackgroundColor.G & ", " & Me.BackgroundColor.B & "); "
        End If
        If Me.BorderStyle = BorderStyleType.Solid Then
            Style &= "border-style: solid; "
            Style &= "border-width: " & Me.BorderWidth & "; "
            Style &= "border-color: rgb(" & Me.BorderColor.R & ", " & Me.BorderColor.G & ", " & Me.BorderColor.B & "); "
        Else
            Style &= "border-style: none; "
        End If
        Return Style
    End Function

End Class

Public Class cLabel
    Inherits cElement

    Public FontName As String = "Arial"
    Public FontSize As Integer = 10
    Public FontColor As System.Drawing.Color = Color.Black
    Public Bold As Boolean = False
    Public Italic As Boolean = False
    Public Underline As Boolean = False

    Public Sub New(ByVal s As String)
        Me._ClassName = "Label"
        Me.Text = s
    End Sub

    Public Function ToXml() As String
        Dim Xml As String = ""
        Xml = "<div id=""" & Me.Id & """ class=""Label"""
        Xml &= " style=""" & Me.StyleString
        Xml &= """>" & Me.Text & "</div>"

        Return Xml
    End Function

End Class

Public Class cImage
    Inherits cElement

    Public Src As String = ""

    Public Property Alt() As String
        Get
            Return Me.Text
        End Get
        Set(ByVal value As String)
            Me.Text = value
        End Set
    End Property

    Public Sub New(ByVal s As String)
        Me.Src = s
        Me._ClassName = "Image"
    End Sub

    Public Function ToXml() As String
        Dim Xml As String = ""
        Xml = "<div id=""" & Me.Id & """ class=""Image"""
        Xml &= " style=""" & Me.StyleString
        Xml &= """>"
        Xml &= "<img src=""" & Me.Src & """ style=""width: 100%; height: 100%"" alt=""" & Me.Text & """/>"
        Xml &= "</div>"
        Return Xml
    End Function

End Class


Public Class cSignatureBox
    Inherits cElement

    Public Src As String = ""

    Public Property Alt()
        Get
            Return Me.Text
        End Get
        Set(ByVal value)
            Me.Text = value
        End Set
    End Property

    Public Sub New(ByVal s As String)
        Me.Src = s
        Me._ClassName = "Signature"
    End Sub

    Public Function ToXml() As String
        Dim Xml As String = ""
        Xml = "<div id=""" & Me.Id & """ class=""Signature"""
        Xml &= " style=""" & Me.StyleString
        Xml &= """>"
        Xml &= "<img src=""" & Me.Src & """ style=""width: 100%; height: 100%"" alt=""" & Me.Text & """/>"
        Xml &= "</div>"
        Return Xml
    End Function

End Class


Public Class cTable
    Inherits cElement

    Public CellPadding As Integer = 0
    Public CellSpacing As Integer = 0
    Public RowLimit As Integer = 0
    Public Source As String = ""

    Public Src As String = ""

    Public Sub New()
        Me._ClassName = "Table"
    End Sub

    Public Function ToXml() As String
        Dim Xml As String = ""
        Xml = "<div id=""" & Me.Id & """ class=""Table"""
        Xml &= " style=""" & Me.StyleString
        Xml &= """>"
        Xml &= "<table"
        Xml &= "</div>"
        Return Xml
    End Function

End Class















Public Class cColumn
    Public Map As String = ""
    Public Format As String = ""
    Public Width As Integer = 50
    Public Align As HorizontalAlignment = HorizontalAlignment.Left
    Public Text As String = ""
End Class

Public Class cCell
    Public Text As String = ""
End Class

Public Class cRow
    Public Cells As New Hashtable
End Class

Public Class cArray

    Dim _Items As New Hashtable
    Dim _InternalCount As Integer = 0

    Public Property Items(ByVal n As Integer) As Object
        Get
            Dim i As Integer = 0
            For Each o As Object In Me._Items
                If n = i Then
                    Return o
                End If
                i += 1
            Next
            Throw New Exception("No item with that index")
        End Get
        Set(ByVal value As Object)
            Dim i As Integer = 0
            For Each o As Object In Me._Items
                If n = i Then
                    o = value
                End If
                i += 1
            Next
            Throw New Exception("No item with that index")
        End Set
    End Property

    Public Sub Add(ByVal Value As Object)
        Me._Items.Add(Me._InternalCount, Value)
        Me._InternalCount += 1
    End Sub

End Class