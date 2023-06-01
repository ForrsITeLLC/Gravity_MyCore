Imports System.Xml
Imports System.Security.Cryptography

Public Class cDocumentPage

    Dim _Xml As New XmlDocument
    Dim _Source As String = ""

    Public Event Cleared()


    Public Property Source() As String
        Get
            Return Me.ToString
        End Get
        Set(ByVal Value As String)
            Me._Xml.LoadXml(Value)
            Me._Source = Value
        End Set
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal Source As String)
        Me._Xml.LoadXml(Source)
        Me._Source = Source
    End Sub

    Public Overrides Function ToString() As String
        Return Me._Xml.SelectSingleNode("/").OuterXml
    End Function

    Public Sub Replace(ByVal OldText As String, ByVal NewText As String)
        Me._Source = Me._Source.Replace(OldText, Me.Escape(NewText))
        Me._Xml.Load(Me._Source)
    End Sub

    Public Function Escape(ByVal Value As String) As String
        Value = Value.Replace("&amp;", "&")
        Value = Value.Replace("&", "&amp;")
        Return Value
    End Function

    Public Sub PopulateTable(ByVal SourceName As String, ByVal Data As DataTable)
        ' Find nodes
        Dim Nodes As XmlNodeList
        Dim Node As XmlNode
        Dim Style As String = ""
        'Style = "font-size: 12px; font-family: Arial; "
        Nodes = Me._Xml.SelectNodes("/html/body/div/table")
        For Each Node In Nodes
            ' If we've found a table with this rouce name
            If Node.SelectSingleNode("@src").InnerText = SourceName Then
                ' Get headers
                Dim Columns(Node.SelectNodes("tr/th").Count - 1) As String
                Dim Formats(Node.SelectNodes("tr/th").Count - 1) As String
                Dim Row As DataRow
                Dim Th As XmlNode
                Dim i As Integer = 0
                ' Loop through xml table to create column array
                For Each Th In Node.SelectNodes("tr/th")
                    Columns(i) = Th.SelectSingleNode("@map").InnerText
                    Try
                        Formats(i) = Th.SelectSingleNode("@format").InnerText
                    Catch ex As Exception
                        Formats(i) = ""
                    End Try
                    If Style.Length > 0 Then
                        Th.Attributes.Append(Me._Xml.CreateAttribute("style")).InnerText = Style
                    End If
                    i += 1
                Next
                ' If there is data in Data variable
                If Data.Rows.Count > 0 Then
                    ' Add Tr For each row
                    For Each Row In Data.Rows
                        Dim Tr As XmlElement = Me._Xml.CreateElement("tr")
                        Tr.Attributes.Append(Me._Xml.CreateAttribute("valign")).InnerText = "top"
                        For i = 0 To Columns.Length - 1
                            Dim Td As XmlElement = Me._Xml.CreateElement("td")
                            Try
                                Td.InnerText = Format(IfNull(Row.Item(Columns(i)), ""), Formats(i))
                            Catch ex As Exception
                                MsgBox("Column did not exist in table.  Template is probably wrong. Here is exact error: " & ex.ToString)
                            End Try
                            If Style.Length > 0 Then
                                Td.Attributes.Append(Me._Xml.CreateAttribute("style")).InnerText = Style
                            End If
                            Tr.AppendChild(Td)
                        Next
                        Node.AppendChild(Tr)
                    Next
                Else
                    Dim Tr As XmlElement = Me._Xml.CreateElement("tr")
                    Dim Td As XmlElement = Me._Xml.CreateElement("td")
                    Dim ColSpan As XmlAttribute = Me._Xml.CreateAttribute("colspan")
                    Td.Attributes.Append(ColSpan).InnerText = Node.SelectNodes("tr/th").Count.ToString
                    Td.InnerText = "None"
                    Tr.AppendChild(Td)
                    Node.AppendChild(Tr)
                End If
            End If
        Next
        Me._Source = Me.ToString
    End Sub

    Private Function IfNull(ByVal Value As Object, Optional ByVal NewVal As String = "") As String
        If Value Is DBNull.Value Then
            Return NewVal
        Else
            Return Value
        End If
    End Function

    Private Function GetElementsByTagName(ByVal Tag As String) As XmlNodeList
        Dim NodeList As XmlNodeList = Me._Xml.SelectNodes("//" & Tag)
        Return NodeList
    End Function

    Private Sub SetValue(ByVal Xpath As String, ByVal Value As String)
        Me._Xml.SelectSingleNode(Xpath).Value = Value
        Me._Source = Me.ToString
    End Sub

    Public Sub MakeImagesExternal()
        Dim Nodes As XmlNodeList
        Dim Node As XmlNode
        Dim strKey As String
        Dim hashImage As New Hashtable
        Dim SrcPath As String
        Dim MD5 As New MD5CryptoServiceProvider
        ' Loop through and replace images
        Nodes = Me._Xml.SelectNodes("//img/@src")
        For Each Node In Nodes
            If Node.InnerText.StartsWith("data:image/gif;base64,") Then
                ' Save to a local image and replace src with path
                Node.InnerText = Node.InnerText.Replace("data:image/gif;base64,", "")
                strKey = BitConverter.ToString(MD5.ComputeHash(System.Text.UnicodeEncoding.ASCII.GetBytes(Node.InnerText)))
                If Not hashImage.ContainsKey(strKey) Then
                    Dim MemStream As New IO.MemoryStream(Convert.FromBase64String(Node.InnerText))
                    Dim Image As New Drawing.Bitmap(MemStream)
                    SrcPath = Environment.GetEnvironmentVariable("TEMP") & "\" & Now.Ticks & "-" & strKey.Substring(0, 8) & ".gif"
                    Image.Save(SrcPath, Drawing.Imaging.ImageFormat.Gif)
                    hashImage.Add(strKey, SrcPath)
                Else
                    SrcPath = hashImage(strKey)
                End If
                Node.InnerText = "file:///" & SrcPath.Replace("\", "/")
            End If
        Next
        Me._Source = Me.ToString
    End Sub

    Public Sub Clear()
        Me._Xml.RemoveAll()
    End Sub

End Class
