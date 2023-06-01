Imports dbAutoTrack

Namespace GravityDocument

    Public Class gPage

        Public Event AddRolloverPage(ByRef ReferringPage As gPage, ByVal Table As gTable)

#Region "Variables"

        Dim _Document As gDocument
        Public Elements As New Collection
        Public Variables As New Hashtable
        Public Id As String = ""
        Dim _LastAddedElement As gElement = Nothing
        Public PageNum As Integer = 0
        Public ContinuedFromPage As gPage

#End Region

#Region "Properties"

        Public Property Document() As gDocument
            Get
                Return Me._Document
            End Get
            Set(ByVal value As gDocument)
                Me._Document = value
            End Set
        End Property

        Public ReadOnly Property LastAddedElement() As Object
            Get
                Return Me._LastAddedElement
            End Get
        End Property

#End Region

#Region "Events"

        Public Event BeforeElementRemoved(ByVal e As gElement)
        Public Event AfterElementRemoved(ByVal Key As String)
        Public Event ElementAdded(ByVal e As gElement)

#End Region

#Region "Methods"

        Public Sub New(ByVal Doc As gDocument, ByVal Id As String)
            Me._Document = Doc
            Me.Id = Id
        End Sub

        Public Sub Clear()
            Me.Elements = New Collection
            Me.Variables = New Hashtable
        End Sub

        Public Sub Render()
            Dim Xml As String = "<html><body><page>" & Me.ToXml(True) & "</page></body></html>"
            Me.Clear()
            ' Output xml
            Try
                Me.LoadXml(Xml)
            Catch
                MsgBox("Error loading: " & Xml)
            End Try
            ' Set page variables
            Me.AddVariable("%p%", Me.PageNum)
            Me.AddVariable("%t%", Me.Document.PageCount)
        End Sub

        Public Sub LoadXml(ByVal Xml As String)
            Dim XmlDoc As New Xml.XmlDocument
            Dim Elements As Xml.XmlNodeList
            Dim PageXml As Xml.XmlNode
            XmlDoc.LoadXml(Xml)
            PageXml = XmlDoc.SelectSingleNode("/html/body/page")
            If PageXml Is Nothing Then
                PageXml = XmlDoc.SelectSingleNode("/html/body/div[@class='Page']")
            End If
            If PageXml Is Nothing Then
                PageXml = XmlDoc.SelectSingleNode("/html/body")
            End If
            Elements = PageXml.SelectNodes("div")
            For Each e As Xml.XmlNode In Elements
                Dim Element As gElement = Me.Document.BuildElement(Me, e)
                Try
                    Me.AddElement(Element)
                Catch ex As Exception
                    Dim Err As String = ex.ToString
                End Try
            Next
        End Sub

        Public Sub PreRender()
            For Each el As gElement In Me.Elements
                If el.IsRollover Then
                    RaiseEvent AddRolloverPage(Me, el.Table)
                End If
            Next
        End Sub

        Public Function ToXml(Optional ByVal InternalImages As Boolean = True, Optional ByVal MoveDown As Integer = 0, Optional ByVal Export As Boolean = False) As String
            ' Do xml
            Dim Xml As String = ""
            For Each el As gElement In Me.Elements
                Dim ElXml As String = el.ToXml(InternalImages, MoveDown, Export)
                Xml &= ElXml & ControlChars.CrLf
            Next
            For Each Var As Variable In Me.Variables.Values
                Xml = Xml.Replace(Var.Name, Me.Document.XmlClean(Var.Value))
            Next
            Return Xml
        End Function

        Public Function GetElement(ByVal i As Integer) As gElement
            Return Me.Elements(i)
        End Function

        Public Function GetElementsOrderedByZIndex(Optional ByVal Order As String = "DESC") As Collection
            Dim Table As New DataTable
            Table.Columns.Add("index")
            Table.Columns.Add("zindex")
            Dim Count As Integer = 1
            For Each el As gElement In Me.Elements
                Dim r As DataRow = Table.NewRow
                r.Item("index") = Count
                r.Item("zindex") = el.ZIndex
                Table.Rows.Add(r)
                Count += 1
            Next
            Dim View As New DataView(Table)
            View.Sort = "zindex " & Order
            Dim OutCollection As New Collection
            For i As Integer = 0 To View.Count - 1
                Dim n As Integer = View.Item(i).Row.Item("index")
                OutCollection.Add(Me.Elements(n))
            Next
            Return OutCollection
        End Function

        Public Function HighestZ() As Integer
            Dim Max As Integer = 0
            For Each el As gElement In Me.Elements
                If el.ZIndex > Max Then
                    Max = el.ZIndex
                End If
            Next
            Return Max
        End Function

        Public Sub IncrementAllZ()
            For Each el As gElement In Me.Elements
                el.ZIndex += 1
            Next
        End Sub

        Public Sub SendToBack(ByVal el As gElement)
            Me.IncrementAllZ()
            el.ZIndex = 0
        End Sub

        Public Sub BringToFront(ByVal el As gElement)
            el.ZIndex = Me.HighestZ + 1
        End Sub

        Public Sub AddElement(ByVal e As gElement)
            Dim Id As String = e.Id.Trim
            Dim Key As String = "element" & Me.Elements.Count
            ' See if key exists
            If Me.Elements.Contains(Key) Then
                Key = "element" & Me.Elements.Count & "_" & Now.Ticks
            End If
            ' Set ID As Key if ID is not valid
            If Id = Nothing Then
                Id = Key
            ElseIf Id.Length = 0 Then
                Id = Key
            End If
            ' Set values
            e.Id = Id
            e.ChangeKey(Key)
            ' Add to collection
            Me.Elements.Add(e, Key)
            ' Record last added
            Me._LastAddedElement = e
            ' Alert the media
            RaiseEvent ElementAdded(e)
        End Sub

        Public Sub RemoveElement(ByVal Key As String)
            RaiseEvent BeforeElementRemoved(Me.Elements(Key))
            Me.Elements.Remove(Key)
            RaiseEvent AfterElementRemoved(Key)
        End Sub

        Public Function GetElementByKey(ByVal Key As String) As gElement
            Return Me.Elements(Key)
        End Function

        Public Function GetElementById(ByVal Id As String) As gElement
            For Each e As gElement In Me.Elements
                If e.Id = Id Then
                    Return e
                End If
            Next
            Return Nothing
        End Function

        Public Function GetTableBySource(ByVal Src As String) As gElement
            For Each e As gElement In Me.Elements
                If e.ClassName = gElement.Classes.Table Then
                    If e.Table.Source = Src Then
                        Return e
                    End If
                End If
            Next
            Return Nothing
        End Function


#End Region

#Region "Variables"

        Public Structure Variable
            Dim Name As String
            Dim Value As String
        End Structure

        Public Sub AddVariable(ByVal Name As String, ByVal Value As String)
            Dim Var As Variable
            Var.Name = Name
            Var.Value = Value
            If Me.Variables.ContainsKey(Name) Then
                Me.Variables(Name) = Var
            Else
                Me.Variables.Add(Name, Var)
            End If
        End Sub

#End Region


        Public Function ToPDFPage() As PDFWriter.Page
            Dim LeftMargin As Double = 25
            Dim TopMargin As Double = 50
            Dim Ratio As Double = 0.8
            Dim Page As New PDFWriter.Page(PDFWriter.PageSize.A4)
            Dim Graphics As PDFWriter.Graphics.PDFGraphics = Page.Graphics
            Dim Col As Collection = Me.GetElementsOrderedByZIndex("ASC")
            Dim Border As PDFWriter.Border
            Dim Fgcolor As PDFWriter.Graphics.RGBColor
            Dim Bgcolor As PDFWriter.Graphics.RGBColor
            Dim FontStyle As System.Drawing.FontStyle
            Dim Font As System.Drawing.Font
            Dim PdfFont As PDFWriter.PDFFont
            Dim Style As PDFWriter.Graphics.TextStyle
            Dim HeaderFontStyle As System.Drawing.FontStyle
            Dim HeaderFont As System.Drawing.Font
            Dim PdfHeaderFont As PDFWriter.PDFFont
            Dim ts As PDFWriter.Graphics.TableStyle
            Dim ths As PDFWriter.Graphics.TableStyle
            Dim th As PDFWriter.Graphics.Row
            Dim tr As PDFWriter.Graphics.Row
            Dim Left As Double
            Dim Top As Double
            For Each Element As gElement In Col
                ' Set up basic properties
                Application.DoEvents()
                ' Border
                Border = New PDFWriter.Border
                Border.LineColor = New PDFWriter.Graphics.RGBColor(Element.BorderColor)
                Border.LineWidth = Element.BorderWidth
                If Element.BorderStyle = gElement.LineStyle.Solid Then
                    Border.LineStyle = PDFWriter.LineStyle.Solid
                ElseIf Element.BorderStyle = gElement.LineStyle.Dotted Then
                    Border.LineStyle = PDFWriter.LineStyle.Dot
                ElseIf Element.BorderStyle = gElement.LineStyle.Dashed Then
                    Border.LineStyle = PDFWriter.LineStyle.Dash
                Else
                    Border = Nothing
                End If
                ' Colors
                Fgcolor = New PDFWriter.Graphics.RGBColor(Element.Fgcolor)
                Bgcolor = New PDFWriter.Graphics.RGBColor(Element.Bgcolor)
                ' Text
                Dim Text As String = Element.Text
                If Text.Contains("%") Then
                    For Each Var As gPage.Variable In Me.Variables.Values
                        Application.DoEvents()
                        Text = Text.Replace(Var.Name, Var.Value)
                    Next
                End If
                ' What kind is it?
                If Element.ClassName = gElement.Classes.Label Then
                    ' LABEL
                    ' Font Style
                    FontStyle = New System.Drawing.FontStyle
                    If Element.FontBold Then
                        FontStyle = FontStyle.Bold
                    End If
                    If Element.FontItalic Then
                        FontStyle = FontStyle Or FontStyle.Italic
                    End If
                    If Element.FontUnderline Then
                        FontStyle = FontStyle Or FontStyle.Underline
                    End If
                    Font = New System.Drawing.Font(Element.FontFamily, Element.FontSize, FontStyle, Drawing.GraphicsUnit.Point)
                    PdfFont = New PDFWriter.PDFFont(Font, False)
                    Style = New PDFWriter.Graphics.TextStyle(PdfFont, Element.FontSize, Fgcolor)
                    ' Free memory
                    FontStyle = Nothing
                    Font = Nothing
                    PdfFont = Nothing
                    ' Alignment
                    Dim Align As PDFWriter.TextAlignment = PDFWriter.TextAlignment.Left
                    If Element.TextAlign = gElement.HorizontalAlignment.Center Then
                        Align = PDFWriter.TextAlignment.Center
                    ElseIf Element.TextAlign = gElement.HorizontalAlignment.Right Then
                        Align = PDFWriter.TextAlignment.Right
                    ElseIf Element.TextAlign = gElement.HorizontalAlignment.Justify Then
                        Align = PDFWriter.TextAlignment.Justify
                    End If
                    ' Padding
                    Dim Padding As Integer = 0
                    If Element.BorderStyle = gElement.LineStyle.Solid Then
                        Padding = 2
                    End If
                    ' Draw label
                    Graphics.DrawTextBox(Element.Left * Ratio + LeftMargin, Element.Top * Ratio + TopMargin, Element.Width * Ratio, Element.Height * Ratio, Text, Style, Align, False, Padding, Bgcolor, Border)
                ElseIf Element.ClassName = gElement.Classes.Image Then
                    ' Draw image
                    Graphics.DrawImage(Element.Image.ToBmp, Element.Left * Ratio + LeftMargin, Element.Top * Ratio + TopMargin, Element.Width * Ratio, Element.Height * Ratio, PDFWriter.SizeMode.Clip, Border)
                ElseIf Element.ClassName = gElement.Classes.Signature Then
                    ' Draw signature
                    Graphics.DrawImage(Element.Image.ToBmp, Element.Left * Ratio + LeftMargin, Element.Top * Ratio + TopMargin, Element.Width * Ratio, Element.Height * Ratio, PDFWriter.SizeMode.Clip, Border)
                ElseIf Element.ClassName = gElement.Classes.Table Then
                    ' Font Style
                    FontStyle = New System.Drawing.FontStyle
                    HeaderFontStyle = New System.Drawing.FontStyle
                    HeaderFontStyle = Drawing.FontStyle.Bold
                    FontStyle = Drawing.FontStyle.Regular
                    Font = New System.Drawing.Font(Element.FontFamily, Element.FontSize, FontStyle, Drawing.GraphicsUnit.Point)
                    HeaderFont = New System.Drawing.Font(Element.FontFamily, Element.FontSize, HeaderFontStyle, Drawing.GraphicsUnit.Point)
                    PdfFont = New PDFWriter.PDFFont(Font, False)
                    PdfHeaderFont = New PDFWriter.PDFFont(HeaderFont, False)
                    ' Create Table
                    Dim Table As New PDFWriter.Graphics.Table
                    If Element.BorderStyle <> gElement.LineStyle.None Then
                        Table.Border = Border
                    End If
                    Table.BackColor = Bgcolor
                    ' Columns
                    For Each c As gTable.Column In Element.Table.Columns
                        ' Create Column
                        Table.Columns.Add(c.Width * Ratio)
                    Next
                    ' Create styles
                    ts = New PDFWriter.Graphics.TableStyle(PdfFont, Element.FontSize, Fgcolor)
                    ths = New PDFWriter.Graphics.TableStyle(PdfHeaderFont, Element.FontSize, Fgcolor)
                    ' Create header row
                    th = New PDFWriter.Graphics.Row(Table, ths)
                    ' Border around it?
                    If Element.BorderStyle <> gElement.LineStyle.None Then
                        th.Border = Border
                    End If
                    ' Create columns and rows
                    For Each c As gTable.Column In Element.Table.Columns
                        Application.DoEvents()
                        ' Create Header
                        Dim ht As String = c.HeaderText
                        If ht.Contains("%") Then
                            For Each Var As gPage.Variable In Me.Variables
                                ht = ht.Replace(Var.Value, Var.Value)
                            Next
                        End If
                        Dim cell As dbAutoTrack.PDFWriter.Graphics.Cell = th.Cells.Add(ht, ths)
                        If c.Align = "right" Then
                            cell.TextAlign = ContentAlignment.MiddleRight
                            cell.ContentAlign = ContentAlignment.MiddleRight
                        ElseIf c.Align = "center" Then
                            cell.TextAlign = ContentAlignment.MiddleCenter
                            cell.ContentAlign = ContentAlignment.MiddleCenter
                        Else
                            cell.TextAlign = ContentAlignment.MiddleLeft
                            cell.ContentAlign = ContentAlignment.MiddleLeft
                        End If
                        If Element.BorderStyle <> gElement.LineStyle.None Then
                            cell.Border = Border
                        End If
                    Next
                    Table.Rows.Add(th)
                    ' Create data rows
                    For Each r As DataRow In Element.Table.Data.Rows
                        Application.DoEvents()
                        tr = New PDFWriter.Graphics.Row(Table, ts)
                        For Each c As gTable.Column In Element.Table.Columns
                            If r.Table.Columns.Contains(c.Key) Then
                                Dim Value As String = IIf(r.Item(c.Key) Is DBNull.Value, "", r.Item(c.Key))
                                Dim cell As dbAutoTrack.PDFWriter.Graphics.Cell = tr.Cells.Add(Format(Value, c.Format), ts)
                                If c.Align = "right" Then
                                    cell.TextAlign = ContentAlignment.MiddleRight
                                    cell.ContentAlign = ContentAlignment.MiddleRight
                                ElseIf c.Align = "center" Then
                                    cell.TextAlign = ContentAlignment.MiddleCenter
                                    cell.ContentAlign = ContentAlignment.MiddleCenter
                                Else
                                    cell.TextAlign = ContentAlignment.MiddleLeft
                                    cell.ContentAlign = ContentAlignment.MiddleLeft
                                End If
                                If Element.BorderStyle <> gElement.LineStyle.None Then
                                    cell.Border = Border
                                End If
                            Else
                                tr.Cells.Add("")
                            End If
                        Next
                        Table.Rows.Add(tr)
                    Next
                    ' Draw Table
                    Left = Element.Left * Ratio + LeftMargin
                    Top = Element.Top * Ratio + TopMargin
                    Graphics.DrawTable(Left, Top, Element.Height, Table)
                End If
            Next
            Return Page
        End Function

    End Class

End Namespace
