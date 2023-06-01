
Namespace Rest

    Public Class ClientKeyValidator

        Public EndPoint As String = "http://www.evware.com/api/rest.php"
        Dim _Validation As String = ""

        Public ReadOnly Property ValidationKey() As String
            Get
                Return Me._Validation
            End Get
        End Property

        Public Function IsValid(ByVal Key As String) As Boolean
            Dim Request As New Request(Me.EndPoint)
            Request.AddArgument("method", "key.isValid")
            Request.AddArgument("key", Key)
            Request.AddArgument("computer", My.Computer.Name)
            Request.Send()
            If Request.LastResponse.Code > 0 Then
                Me._Validation = Request.LastResponse.Content
                Return True
            Else
                Return False
            End If
        End Function

    End Class

    Public Class Update

        Public EndPoint As String = "http://www.evware.com/api/rest.php"

        Public DateReleased As Date
        Public Description As String
        Public Priority As Integer
        Public Url As String
        Public LastResponse As Request.Response

        Public Function GetLatest(ByVal Key As String) As Boolean
            Dim Request As New Request(Me.EndPoint)
            Request.AddArgument("method", "update.getLatest")
            Request.AddArgument("key", Key)
            Request.Send()
            Me.LastResponse = Request.LastResponse
            If Request.LastResponse.Code > 0 Then
                Dim Doc As New Xml.XmlDocument
                Doc.Load(Request.LastResponseText)
                Me.DateReleased = Doc.SelectSingleNode("/response/content/date").InnerText
                Me.Description = Doc.SelectSingleNode("/response/content/description").InnerText
                Me.Priority = Doc.SelectSingleNode("/response/content/priority").InnerText
                Me.Url = Doc.SelectSingleNode("/response/content/url").InnerText
                Return True
            Else
                Return False
            End If
        End Function



    End Class

    Public Class Request

        Dim _Response As String = ""

        Public EndPoint As String = ""
        Dim Args As New Collection
        Public Method As String = "GET"
        Public LastResponse As Response

        Public Class Response
            Public Code As Integer
            Public Msg As String
            Public Content As String
            Public Sub New(ByVal n As Integer, ByVal m As String, ByVal c As String)
                Me.Code = n
                Me.Msg = m
                Me.Content = c
            End Sub
        End Class

        Public Class Param
            Public Name As String
            Public Value As String
            Public Sub New(ByVal n As String, ByVal v As String)
                Me.Name = n
                Me.Value = v
            End Sub
        End Class

        Public ReadOnly Property QueryString() As String
            Get
                Dim Query As String = ""
                Dim Count As Integer = 0
                For Each a As Param In Args
                    If Count > 0 Then
                        Query &= "&"
                    End If
                    Query &= a.Name & "=" & a.Value
                    Count += 1
                Next
                Return Query
            End Get
        End Property

        Public ReadOnly Property Url() As String
            Get
                If Me.Method = "GET" Then
                    Return Me.EndPoint & "?" & Me.QueryString
                Else
                    Return Me.EndPoint
                End If
            End Get
        End Property

        Public ReadOnly Property LastResponseText() As String
            Get
                Return Me._Response
            End Get
        End Property

        Public Sub New(ByVal EndPoint As String)
            Me.EndPoint = EndPoint
        End Sub

        Public Sub AddArgument(ByVal Name As String, ByVal Value As String)
            Me.Args.Add(New Request.Param(Name, Value))
        End Sub

        Public Sub Send()
            ' Create Request
            Dim Request As System.Net.HttpWebRequest = System.Net.WebRequest.Create(Me.Url)
            Request.Method = Me.Method
            Request.ContentType = "application/x-www-form-urlencoded"
            Request.UserAgent = "Gravity"
            If Me.Method = "POST" Then
                ' Create Request Content
                Dim Enc As New System.Text.UTF8Encoding
                Dim Data() As Byte = Enc.GetBytes(Me.QueryString)
                ' Add Content to Request
                Request.ContentLength = Data.Length
                Dim Stream As System.IO.Stream = Request.GetRequestStream()
                Stream.Write(Data, 0, Data.Length)
                Stream.Close()
            End If
            ' Get Response
            Dim Response As System.Net.HttpWebResponse = Request.GetResponse
            Dim Reader As New System.IO.StreamReader(Response.GetResponseStream, System.Text.Encoding.ASCII)
            Me._Response = Reader.ReadToEnd
            Reader.Close()
            Me.ProcessResponse()
        End Sub

        Private Sub ProcessResponse()
            Dim Doc As New Xml.XmlDocument
            Doc.LoadXml(Me._Response)
            Dim Code As Integer = Doc.SelectSingleNode("/response/code").InnerText
            Dim Msg As String = Doc.SelectSingleNode("/response/message").InnerText
            Dim Content As String = Doc.SelectSingleNode("/response/content").InnerText
            Me.LastResponse = New Response(Code, Msg, Content)
        End Sub


    End Class


End Namespace
