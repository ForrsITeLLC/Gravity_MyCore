Imports System.Xml
Imports dbAutoTrack

Namespace GravityDocument

    Public Class gDocument

        Public Pages As New Collection
        Public Title As String = ""
        Public Author As String = ""
        Public ReferenceID As String = ""
        Public FormTypeID As Integer = 0
        Public CanEdit As Boolean = True

        Public PageHeight As Integer = 0

        Public LocalImages As New Hashtable

        Dim _RolloverPageXml As String = ""

        Public Property RolloverPageXml() As String
            Get
                If Me._RolloverPageXml.Length = 0 Then
                    Dim Xml As String = ""
                    Xml &= "<html>"
                    Xml &= "<head><title>" & Me.XmlClean(Me.Title) & "</title></head>"
                    Xml &= "<body>"
                    Xml &= "</body></html>"
                    Return Xml
                Else
                    Return Me._RolloverPageXml
                End If
            End Get
            Set(ByVal value As String)
                Me._RolloverPageXml = value
            End Set
        End Property


        Public Property FormType() As FormTypes
            Get
                Return Me.FormTypeID
            End Get
            Set(ByVal value As FormTypes)
                Me.FormTypeID = CInt(value)
            End Set
        End Property

        Public Enum FormTypes
            Other = 0
            ServiceOrder = 1
            SalesOrder = 2
            Quote = 3
            CalCert = 4
            Invoice = 5
            RentalOrder = 6
            PurchaseOrder = 7
            Report = 8
            InteractionLetter = 9
            PaymentReceipt = 10
            CustomerStatement = 11
        End Enum

        Public Sub New(ByVal PgHeight As Integer)
            Me.PageHeight = PgHeight
        End Sub

        Public ReadOnly Property PageCount() As Integer
            Get
                Return Me.Pages.Count
            End Get
        End Property

        Public ReadOnly Property GetPage(ByVal i As Integer) As gPage
            Get
                If i <= Me.Pages.Count Then
                    Return Me.Pages(i)
                Else
                    Return Nothing
                End If
            End Get
        End Property

        Public Function AddPageFromXml(ByVal Xml As String) As gPage
            Dim Page As New gPage(Me, "page" & Me.PageCount + 1)
            Page.LoadXml(Xml)
            Me.Pages.Add(Page)
            Return Page
        End Function

        Public Function AddPage(ByVal pg As gPage) As gPage
            pg.Id = "page" & (Me.PageCount + 1)
            pg.Document = Me
            pg.PageNum = Me.PageCount + 1
            Me.Pages.Add(pg)
            Return pg
        End Function

        Public Sub AddPage(ByVal doc As gDocument, Optional ByVal n As Integer = 1)
            doc.PreRenderPages()
            For Each pg As gPage In doc.Pages
                Me.AddPage(pg)
            Next
        End Sub

        Public Function InsertPage(ByVal Xml As String, ByVal InsertAfter As gPage) As gPage
            Dim Page As New gPage(Me, "new page")
            Page.LoadXml(Xml)
            Return Me.InsertPage(Page, InsertAfter)
        End Function

        Public Function InsertPage(ByVal pg As gPage, ByVal InsertAfter As gPage) As gPage
            If InsertAfter IsNot Nothing Then
                ' We are going to rebuild the array
                Dim Count As Integer = 1
                Dim NewPageArray As New Collection
                For Each ExistingPage As gPage In Me.Pages
                    ' Add existing
                    ExistingPage.Id = "page" & Count
                    ExistingPage.PageNum = Count
                    NewPageArray.Add(ExistingPage)
                    Count += 1
                    ' If this is where it should inserted
                    If InsertAfter Is ExistingPage Then
                        pg.Id = "page" & Count
                        pg.Document = Me
                        pg.PageNum = Count
                        NewPageArray.Add(pg)
                        Count += 1
                    End If
                Next
                Me.Pages = NewPageArray
            Else
                ' Insert at end
                Dim Count As Integer = Me.Pages.Count + 1
                pg.Id = "page" & Count
                pg.Document = Me
                pg.PageNum = Count
                pg = Me.AddPage(pg)
            End If
            Return pg
        End Function

        Public Function InsertBefore(ByVal Xml As String, ByVal InsertBeforePg As gPage) As gPage
            Dim Page As New gPage(Me, "new page")
            Page.LoadXml(Xml)
            Return Me.InsertBefore(Page, InsertBeforePg)
        End Function

        Public Function InsertBefore(ByVal pg As gPage, ByVal InsertBeforePg As gPage) As gPage
            If InsertBeforePg IsNot Nothing Then
                ' We are going to rebuild the array
                Dim Count As Integer = 1
                Dim NewPageArray As New Collection
                For Each ExistingPage As gPage In Me.Pages
                    ' If this is where it should inserted
                    If InsertBeforePg Is ExistingPage Then
                        pg.Id = "page" & Count
                        pg.Document = Me
                        pg.PageNum = Count
                        NewPageArray.Add(pg)
                        Count += 1
                    End If
                    ' Add existing
                    ExistingPage.Id = "page" & Count
                    ExistingPage.PageNum = Count
                    NewPageArray.Add(ExistingPage)
                    Count += 1
                Next
                Me.Pages = NewPageArray
            Else
                ' Insert at end
                Dim Count As Integer = Me.Pages.Count + 1
                pg.Id = "page" & Count
                pg.Document = Me
                pg.PageNum = Count
                pg = Me.AddPage(pg)
            End If
            Return pg
        End Function

        Public Sub AddElementFromXml(ByVal Page As gPage, ByVal Xml As String, Optional ByVal X As Integer = Nothing, Optional ByVal Y As Integer = Nothing)
            Dim XmlDoc As New XmlDocument
            XmlDoc.LoadXml(Xml)
            Dim Node As XmlNode = XmlDoc.SelectSingleNode("/div")
            If Node IsNot Nothing Then
                Dim Element As gElement = Me.BuildElement(Page, Node)
                If X <> Nothing Then
                    Element.Left = X
                End If
                If Y <> Nothing Then
                    Element.Top = Y
                End If
                Page.AddElement(Element)
            Else
                Throw New Exception("XML node /div not found")
            End If
        End Sub

        Public Sub LoadXml(ByVal Xml As String)
            Dim XmlDoc As New Xml.XmlDocument
            XmlDoc.LoadXml(Xml)
            Dim i As Integer = 1
            Try
                Me.Title = XmlDoc.SelectSingleNode("/html/head/title").InnerText
            Catch ex As Exception
                ' Nothing
            End Try
            Dim XmlPages As XmlNodeList
            XmlPages = XmlDoc.SelectNodes("/html/body/div[@class='Page']")
            If XmlPages.Count = 0 Then
                XmlPages = XmlDoc.SelectNodes("/html/body/page")
            End If
            If XmlPages.Count > 0 Then
                For Each p As Xml.XmlNode In XmlPages
                    Application.DoEvents()
                    Dim Page As New gPage(Me, "page" & i)
                    Dim Elements As Xml.XmlNodeList = p.SelectNodes("div")
                    For Each e As Xml.XmlNode In Elements
                        Application.DoEvents()
                        Dim Element As gElement = Me.BuildElement(Page, e)
                        ' If not page 1, move up as if it were the only page
                        If i > 1 Then
                            Element.Top = Element.Top - Me.PageHeight * (i - 1)
                        End If
                        ' Add to page
                        Page.AddElement(Element)
                    Next
                    Me.Pages.Add(Page)
                    i += 1
                Next
            Else
                Dim p As Xml.XmlNode = XmlDoc.SelectSingleNode("/html/body")
                If p IsNot Nothing Then
                    Dim Page As New gPage(Me, "page1")
                    Dim Elements As Xml.XmlNodeList = p.SelectNodes("div")
                    For Each e As Xml.XmlNode In Elements
                        Application.DoEvents()
                        Dim Element As gElement = Me.BuildElement(Page, e)
                        Try
                            Page.AddElement(Element)
                        Catch ex As Exception
                            Dim Err As String = ex.ToString
                        End Try
                    Next
                    Me.Pages.Add(Page)
                End If
            End If
        End Sub

        Public Function BuildElement(ByVal Page As gPage, ByVal Node As Xml.XmlNode) As gElement
            ' Create element
            Dim Element As gElement
            ' Which class?
            Select Case Node.SelectSingleNode("@class").Value
                Case "Image"
                    Element = New gElement(Page, "", gElement.Classes.Image)
                    Element = Me.FormatImage(Element, Node)
                Case "Signature"
                    Element = New gElement(Page, "", gElement.Classes.Signature)
                    Element = Me.FormatSignature(Element, Node)
                Case "Table"
                    Element = New gElement(Page, "", gElement.Classes.Table)
                    Element = Me.FormatTable(Element, Node)
                Case Else
                    Element = New gElement(Page, "", gElement.Classes.Label)
                    Element = Me.FormatLabel(Element, Node)
            End Select
            ' Save original xml
            Element.OriginalSource = Node.OuterXml
            ' Set Id
            Dim IdNode As XmlNode = Node.SelectSingleNode("@id")
            If IdNode IsNot Nothing Then
                Element.Id = IdNode.Value
            End If
            ' Extract CSS Settings
            Dim Style As Hashtable = Me.ExtractStyle(Node.SelectSingleNode("@style").Value)
            Element.Height = Style.Item("height")
            Element.Width = Style.Item("width")
            Element.Left = Style.Item("left")
            Element.Top = Style.Item("top")
            If Style.ContainsKey("font-size") Then
                Element.FontSize = Style.Item("font-size")
            End If
            Element.FontBold = IIf(Style.Item("font-weight") = "bold", True, False)
            Element.FontItalic = IIf(Style.Item("font-style") = "italic", True, False)
            If Style.ContainsKey("background-color") Then
                Element.Bgcolor = Style.Item("background-color")
            Else
                Element.Bgcolor = Drawing.Color.Transparent
            End If
            If Style.ContainsKey("color") Then
                Element.Fgcolor = Style.Item("color")
            Else
                Element.Fgcolor = Drawing.Color.Black
            End If
            If Style.Contains("z-index") Then
                Element.ZIndex = Style.Item("z-index")
            End If
            If Style.Contains("text-align") Then
                If Style.Item("text-align").ToString.ToLower = "center" Then
                    Element.TextAlign = gElement.HorizontalAlignment.Center
                ElseIf Style.Item("text-align").ToString.ToLower = "justify" Then
                    Element.TextAlign = gElement.HorizontalAlignment.Justify
                ElseIf Style.Item("text-align").ToString.ToLower = "right" Then
                    Element.TextAlign = gElement.HorizontalAlignment.Right
                ElseIf Style.Item("text-align").ToString.ToLower = "center" Then
                    Element.TextAlign = gElement.HorizontalAlignment.Left
                End If
            End If
            ' Border
            If Style.Item("border-style") = "solid" Then
                Element.BorderStyle = gElement.LineStyle.Solid
            Else
                Element.BorderStyle = gElement.LineStyle.None
            End If
            Element.BorderWidth = Style.Item("border-width")
            If Style.ContainsKey("border-color") Then
                Element.BorderColor = Style.Item("border-color")
            Else
                Element.BorderColor = Drawing.Color.Black
            End If
            ' Return it
            Return Element
        End Function

        Private Function FormatLabel(ByVal Element As gElement, ByVal Node As Xml.XmlNode) As gElement
            Element.Text = Me.TextClean(Node.InnerText)
            Return Element
        End Function

        Private Function FormatImage(ByVal Element As gElement, ByVal Node As Xml.XmlNode) As gElement
            Dim Img As XmlNode = Node.SelectSingleNode("img")
            Dim Src As String = Img.SelectSingleNode("@src").Value
            If Img.SelectSingleNode("alt") IsNot Nothing Then
                Element.Text = Me.TextClean(Img.SelectSingleNode("alt").Value)
            End If
            If Src.StartsWith("data:") Then
                Element.Image.LoadFromDataUrl(Src)
            End If
            Return Element
        End Function


        Private Function FormatTable(ByVal Element As gElement, ByVal Node As Xml.XmlNode) As gElement
            ' Data
            Try
                Element.Table.Source = Node.FirstChild.SelectSingleNode("@src").InnerText
            Catch
                Element.Table.Source = ""
            End Try
            Try
                Element.Table.RowsPerPage = Node.FirstChild.SelectSingleNode("@rowlimit").InnerText
            Catch
                Element.Table.RowsPerPage = 0   ' Unlimited
            End Try
            Try
                Element.Table.CellPadding = Node.FirstChild.SelectSingleNode("@cellpadding").InnerText
            Catch
                ' Nothing
            End Try
            Try
                Element.Table.CellSpacing = Node.FirstChild.SelectSingleNode("@cellspacing").InnerText
            Catch
                ' Nothing
            End Try
            ' Get tbody
            Dim Tbody As XmlElement = Node.FirstChild
            ' Columns
            For Each Th As XmlNode In Tbody.FirstChild.SelectNodes("th")
                Dim col As New gTable.Column(Element.Table)
                If Th.SelectSingleNode("@map") IsNot Nothing Then
                    col.Key = Th.SelectSingleNode("@map").InnerText
                End If
                If Th.SelectSingleNode("@format") IsNot Nothing Then
                    col.Format = Th.SelectSingleNode("@format").InnerText
                End If
                If Th.SelectSingleNode("@align") IsNot Nothing Then
                    col.Align = Th.SelectSingleNode("@align").InnerText
                Else
                    col.Align = "left"
                End If
                If Th.SelectSingleNode("@width") IsNot Nothing Then
                    col.Width = Th.SelectSingleNode("@width").InnerText
                End If
                col.HeaderText = Th.InnerText
                Element.Table.AddColumn(col)
                Element.Table.Data.Columns.Add(Th.SelectSingleNode("@map").InnerText)
            Next
            ' DATA
            ' Get rows
            Dim Rows As XmlNodeList = Tbody.SelectNodes("tr")
            ' Ignore first row because it is for the column definition
            If Rows.Count > 1 Then
                ' Loop through each row starting with the second
                For i As Integer = 1 To Rows.Count - 1
                    Dim tr As DataRow = Element.Table.Data.NewRow
                    ' Loop through each cell in this row
                    Dim j As Integer = 1
                    For Each td As XmlNode In Rows(i).SelectNodes("td")
                        ' Add cell value, j represents column count
                        tr.Item(Element.Table.Columns(j).Key) = Format(td.InnerText, Element.Table.Columns(j).Format)
                        j += 1
                    Next
                    ' Add row to table
                    Element.Table.Data.Rows.Add(tr)
                Next
            End If
            Return Element
        End Function

        Private Function FormatSignature(ByVal Element As gElement, ByVal Node As Xml.XmlNode) As gElement
            Dim Img As XmlNode = Node.SelectSingleNode("img")
            Dim Src As String = Img.SelectSingleNode("@src").Value
            ' Data
            If Src.StartsWith("data:") Then
                Element.Image.LoadFromDataUrl(Src)
            End If
            Return Element
        End Function

        Private Function ExtractStyle(ByVal Style As String) As Hashtable
            Dim Table As New Hashtable
            Dim SplitStyle As String() = Style.Split(";")
            Dim Attrib(1) As String
            Dim StyleAttrib As String
            For Each StyleAttrib In SplitStyle
                If StyleAttrib.Trim.Length > 0 Then
                    Attrib = StyleAttrib.Split(":")
                    Dim Name As String = Attrib(0).Trim
                    Dim Value As Object = Nothing
                    Select Case Name
                        Case "font-size", "border-width", "left", "top", "padding", "margin", "width", "height"
                            Attrib(1) = Attrib(1).Replace("pt", "")
                            Attrib(1) = Attrib(1).Replace("em", "")
                            Attrib(1) = Attrib(1).Replace("px", "")
                            Attrib(1) = Attrib(1).Replace("cm", "")
                            Attrib(1) = Attrib(1).Replace("ex", "")
                            Attrib(1) = Attrib(1).Replace("mm", "")
                            Attrib(1) = Attrib(1).Replace("in", "")
                            Value = Attrib(1).Trim
                        Case "color", "background-color", "border-color"
                            Dim SplitColor(2) As String
                            Attrib(1) = Attrib(1).Replace("rgb(", "")
                            Attrib(1) = Attrib(1).Replace(")", "")
                            SplitColor = Attrib(1).Split(",")
                            Value = System.Drawing.Color.FromArgb(SplitColor(0).Trim, SplitColor(1).Trim, SplitColor(2).Trim)
                        Case Else
                            Value = Attrib(1).Trim
                    End Select
                    Table.Add(Name, Value)
                End If
            Next
            Return Table
        End Function

        Public Function ExtractSingleStyle(ByVal Node As XmlNode, ByVal StyleProperty As String) As String
            Dim Style As String() = Node.SelectSingleNode("@style").InnerText.Split(";")
            For i As Integer = 0 To Style.Length - 1
                If Style(i).Split(":")(0).Trim = StyleProperty Then
                    Return Style(i).Split(":")(1).Trim
                End If
            Next
            Return Nothing
        End Function

        Public Function ToXml(ByVal InternalImages As Boolean, Optional ByVal ReplaceVariables As Boolean = False, Optional ByVal Export As Boolean = False) As String
            Dim Xml As String = ""
            Xml &= "<html>"
            Xml &= "<head><title>" & Me.XmlClean(Me.Title) & "</title></head>"
            Xml &= "<body>"
            Xml &= Me.RenderPageXml(InternalImages, ReplaceVariables, Export)
            Xml &= "</body></html>"
            Me.LocalImages.Clear()
            Return Xml
        End Function

        Private Sub PreRenderPages()
            For Each Page As gPage In Me.Pages
                AddHandler Page.AddRolloverPage, AddressOf Me.AddRolloverPage
                Page.PreRender()
            Next
        End Sub

        Public Sub Render()
            Me.PreRenderPages()
            For Each Page As MyCore.GravityDocument.gPage In Me.Pages
                Page.Render()
            Next
        End Sub

        Private Function RenderPageXml(ByVal InternalImages As Boolean, Optional ByVal ReplaceVariables As Boolean = False, Optional ByVal Export As Boolean = False) As String
            Dim Xml As String = ""
            Dim i As Integer = 0
            For Each Page As gPage In Me.Pages
                Page.PageNum = i + 1
                Xml &= "<page id=""" & Page.Id & """>" & ControlChars.CrLf
                Xml &= Page.ToXml(InternalImages, Me.PageHeight * i, Export)
                Xml &= "</page>" & ControlChars.CrLf
                i += 1
            Next
            Return Xml
        End Function

        Private Sub AddRolloverPage(ByRef ContinuedFrom As gPage, ByVal Table As gTable)
            ' Create page
            Dim Page As New gPage(Me, "temp")
            Page.ContinuedFromPage = ContinuedFrom
            ' Build page
            Page.LoadXml(Me.RolloverPageXml)
            ' Create place holder
            Dim Rollover As gElement
            ' Get table if it exists
            Rollover = Page.GetElementById("RolloverTable")
            ' If it did not exist, create it
            If Rollover Is Nothing Then
                Rollover = New gElement(Page, "RolloverTable", gElement.Classes.Table)
                Rollover.Left = 20
                Rollover.Top = 20
                Rollover.Width = 600
                Rollover.Height = 900
                Rollover.ZIndex = 100
                Page.AddElement(Rollover)
            End If
            ' New Table... with first x elements removed
            Dim ClonedTable As DataTable = Table.Data.Clone
            Dim StartAt As Integer = Table.RowsPerPage - 1
            For i As Integer = StartAt To Table.Data.Rows.Count - 1
                ClonedTable.ImportRow(Table.Data.Rows(i))
            Next
            If Rollover.ClassName = gElement.Classes.Table Then
                ' Reset columns to match table
                Rollover.Table.ClearColumns()
                ' Loop through columns and add
                For Each Col As gTable.Column In Table.Columns
                    Rollover.Table.AddColumn(Col, Col.Key)
                Next
                ' Add datasource
                Rollover.Table.Data = ClonedTable
            End If
            ' Insert page after the one its continued from
            Me.InsertPage(Page, ContinuedFrom)
            ' Pre render the rollover page in case it needs a rollover page itself!
            Page.PreRender()
        End Sub

        Public Function TextClean(ByVal Text As String) As String
            ' Clean up line breaks
            If Text.Contains(ControlChars.Lf) And Not Text.Contains(ControlChars.CrLf) Then
                Text = Text.Replace(ControlChars.Lf, ControlChars.CrLf)
            ElseIf Text.Contains(ControlChars.Cr) And Not Text.Contains(ControlChars.CrLf) Then
                Text = Text.Replace(ControlChars.Cr, ControlChars.CrLf)
            End If
            ' Trim trailing whitespace
            Text = Text.TrimEnd
            Return Text
        End Function


        Public Function XmlClean(ByVal Value As Object) As String
            Dim Text As String = ""
            ' Nulls
            If Value Is DBNull.Value Then
                Text = ""
            ElseIf Value Is Nothing Then
                Text = ""
            Else
                Text = Value.ToString
            End If
            ' Escape ampresands
            Text = Text.Replace("&amp;", "&")   ' We have to do this to avoid &amp;amp;
            Text = Text.Replace("&", "&amp;")
            ' Escape greater than
            Text = Text.Replace(">", "&gt;")
            ' Escape less than
            Text = Text.Replace("<", "&lt;")
            ' Put in line breaks
            'If Text.Contains(ControlChars.CrLf) Then
            '    Text = Text.Replace(ControlChars.CrLf, "<br />" & ControlChars.CrLf)
            'ElseIf Text.Contains(ControlChars.Lf) Then
            '    Text = Text.Replace(ControlChars.Lf, "<br />" & ControlChars.CrLf)
            'ElseIf Text.Contains(ControlChars.Cr) Then
            '    Text = Text.Replace(ControlChars.Cr, "<br />" & ControlChars.CrLf)
            'End If
            Text = Text.Replace(ControlChars.CrLf, "<br />")
            Text = Text.Replace(ControlChars.Cr, "")
            Text = Text.Replace(ControlChars.Lf, "")
            Text = Text.Replace("<br />", "<br />" & ControlChars.CrLf)
            ' Return cleaned text
            Return Text
        End Function

        Public Sub SaveToEml(ByVal FileName As String)
            ' Render it
            Me.Render()
            ' Save as  html
            Dim htmlfile As String = FileName & ".html"
            My.Computer.FileSystem.WriteAllText(htmlfile, Me.ToXml(False, True, True), False)
            ' Get MHT Object
            Dim mht As New Chilkat.Mht
            mht.UnlockComponent("anything")
            mht.UseCids = True
            ' Write to string
            Dim MhtText As String = mht.HtmlToEML(My.Computer.FileSystem.ReadAllText(htmlfile))
            ' Save to file
            Dim sw As New IO.StreamWriter(FileName, False, System.Text.Encoding.ASCII)
            sw.Write(MhtText)
            sw.Close()
            ' Delete html
            My.Computer.FileSystem.DeleteFile(htmlfile)
        End Sub

        Public Sub SaveToMht(ByVal FileName As String)
            ' Render it
            Me.Render()
            ' Save as  html
            Dim htmlfile As String = FileName & ".html"
            My.Computer.FileSystem.WriteAllText(htmlfile, Me.ToXml(False, True, True), False)
            ' Get MHT Object
            Dim mht As New Chilkat.Mht
            mht.UnlockComponent("anything")
            mht.UseCids = True
            ' Write to string
            Dim MhtText As String = mht.HtmlToMHT(My.Computer.FileSystem.ReadAllText(htmlfile))
            ' Save to file
            Dim sw As New IO.StreamWriter(FileName, False, System.Text.Encoding.ASCII)
            sw.Write(MhtText)
            sw.Close()
            ' Delete html
            My.Computer.FileSystem.DeleteFile(htmlfile)
        End Sub

        Public Sub SaveToPdf(ByVal FileName As String, Optional ByVal StartingPage As Integer = 1, Optional ByVal NumberOfPages As Integer = 0)
            ' Render it
            Me.Render()
            ' Create pdf
            Dim Doc As New PDFWriter.Document
            Doc.Title = Me.Title
            Doc.Author = ""
            Doc.Creator = "EVware Gravity 2"
            Dim PageNum As Integer = 1
            For Each objPage As gPage In Me.Pages
                If PageNum >= StartingPage And (NumberOfPages = 0 Or PageNum <= StartingPage + NumberOfPages) Then
                    Try
                        Doc.Pages.Add(objPage.ToPDFPage)
                    Catch ex As Exception
                        Dim Err As New MyCore.Gravity.ErrorBox("Error creating page at page # " & PageNum & ". The rest of the document could not be completed.", "PDF Error", ex.ToString)
                        Exit For
                    End Try
                End If
                PageNum += 1
            Next
            ' Save File
            Dim Stream As New System.IO.FileStream(FileName, IO.FileMode.Create, IO.FileAccess.Write)
            Try
                Doc.Generate(Stream)
            Catch ex As Exception
                Throw New Exception(ex.ToString)
            Finally
                Stream.Flush()
                Stream.Close()
            End Try
        End Sub

        Public Function SaveToTemp(ByVal Format As String, Optional ByVal ThenOpen As Boolean = False) As String
            ' Adjust format
            Format = Format.ToLower
            If Format = "htm" Then
                Format = "html"
            End If
            ' Make file
            Dim FilePath As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\gravity." & Now.Ticks & "." & Format
            If Format = "pdf" Then
                Me.SaveToPdf(FilePath)
            ElseIf Format = "html" Then
                ' Render it
                Me.Render()
                ' Write it
                My.Computer.FileSystem.WriteAllText(FilePath, Me.ToXml(False, False, True), False)
            ElseIf Format = "mht" Then
                Me.SaveToMht(FilePath)
            ElseIf Format = "eml" Then
                Me.SaveToEml(FilePath)
            Else
                Dim Err As New MyCore.Gravity.ErrorBox("Unknown save format: " & Format, "Error")
                Return Nothing
            End If
            If My.Computer.FileSystem.FileExists(FilePath) Then
                If ThenOpen Then
                    Try
                        System.Diagnostics.Process.Start(FilePath)
                    Catch ex As Exception
                        Dim Err As New MyCore.Gravity.ErrorBox("Open command failed.", "Error", "File: " & FilePath & " -- " & ex.ToString)
                    End Try
                End If
                Return FilePath
            Else
                Return Nothing
            End If
        End Function

        Public Sub Print()
            Dim FilePath As String = Me.SaveToTemp("html")
            Dim Print As New MyCore.frmPrint(FilePath)
            Print.ShowDialog()
        End Sub

        Public Sub Fax(ByVal FaxServer As String, ByVal FaxNumber As String, ByVal ReceipientName As String)
            Dim Server As New FAXCOMLib.FaxServer
            Dim Document As FAXCOMLib.FaxDoc
            Server.Connect(FaxServer)
            Document = Server.CreateDocument(Me.SaveToTemp("pdf"))
            Document.SendCoverpage = 0
            Document.FaxNumber = FaxNumber
            Document.RecipientName = ReceipientName
            Document.DisplayName = ReceipientName
            Document.Send()
            Server.Disconnect()
        End Sub

        Public Sub Email(ByRef ParentWin As MyCore.Plugins.Host, ByVal ToEmail As String, ByVal ToName As String, ByVal Subject As String, ByVal Body As String)
            Dim Com As New MyCore.Utility.Communication(ParentWin.SettingsGlobal, ParentWin.CurrentUser)
            Dim Message As MyCore.Email.Message = Com.CreateNewEmail(ToEmail, ToName)
            Message.Subject = Subject
            Message.Body = Body
            Message.AddAttachment(Me.SaveToTemp("pdf"))
            Message.Send()
        End Sub

    End Class

End Namespace
