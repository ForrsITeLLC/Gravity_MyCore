Imports System.Windows.Forms
Imports System.Xml
Imports System.Drawing
Imports System.Security.Cryptography

Namespace Reports

    Public Class Standards

        Public Shared Function GetAll(ByRef db As MyCore.Data.EasySql, Optional ByVal Since As DateTime = Nothing) As DataTable
            Dim Sql As String = ""
            Sql &= "SELECT e.*,"
            Sql &= " c.test_no, c.date_tested, c.date_expires, c.id AS cert_id, s.name AS location_name"
            Sql &= " FROM standards_equipment e"
            Sql &= " LEFT OUTER JOIN station s ON e.station_id=s.id"
            Sql &= " LEFT OUTER JOIN (SELECT stc.asset_no, c.* FROM standards_to_cert stc LEFT JOIN standards_certification c ON stc.test_no=c.test_no) c"
            Sql &= " ON e.asset_no=c.asset_no"
            If Since <> Nothing Then
                Sql &= " WHERE e.date_last_updated > " & db.Escape(Since)
            End If
            Sql &= " ORDER BY e.asset_no, date_expires DESC"
            Dim Table As DataTable = db.GetAll(Sql)
            Dim Standards As DataTable = Table.Clone
            Dim AssetHash As New Hashtable
            For Each Row As DataRow In Table.Rows
                If Not AssetHash.ContainsKey(Row.Item("asset_no")) Then
                    Dim nr As DataRow = Standards.NewRow
                    For i As Integer = 0 To Row.ItemArray.Length - 1
                        nr.Item(i) = Row.Item(i)
                    Next
                    Standards.Rows.Add(nr)
                    AssetHash.Add(Row.Item("asset_no"), "")
                End If
            Next
            Return Standards
        End Function

        Public Shared Function GetBySerial(ByRef db As MyCore.Data.EasySql, ByVal Serial As String) As DataTable
            Dim Sql As String = ""
            Sql &= "SELECT e.*,"
            Sql &= " c.test_no, c.date_tested, c.date_expires, c.id AS cert_id, s.name AS location_name"
            Sql &= " FROM standards_equipment e"
            Sql &= " LEFT OUTER JOIN station s ON e.station_id=s.id"
            Sql &= " LEFT OUTER JOIN (SELECT stc.asset_no, c.* FROM standards_to_cert stc LEFT JOIN standards_certification c ON stc.test_no=c.test_no) c"
            Sql &= " ON e.asset_no=c.asset_no"
            Sql &= " WHERE e.serial_no LIKE " & db.Escape("%" & Serial & "%")
            Sql &= " ORDER BY e.asset_no, date_expires DESC"
            Dim Table As DataTable = db.GetAll(Sql)
            Dim Standards As DataTable = Table.Clone
            Dim AssetHash As New Hashtable
            For Each Row As DataRow In Table.Rows
                If Not AssetHash.ContainsKey(Row.Item("asset_no")) Then
                    Dim nr As DataRow = Standards.NewRow
                    For i As Integer = 0 To Row.ItemArray.Length - 1
                        nr.Item(i) = Row.Item(i)
                    Next
                    Standards.Rows.Add(nr)
                    AssetHash.Add(Row.Item("asset_no"), "")
                End If
            Next
            Return Standards
        End Function

    End Class

    Public Class TemplateViewer

        Public Event SelectedIndexChanged(ByVal Container As Panel)
        Public Event SelectedElement_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
        Public Event SelectedElement_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
        Public Event SelectedElement_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
        Public Event SelectedElement_DoubleClick(ByVal sender As Object, ByVal e As EventArgs)
        Public Event SelectedElement_KeyPress(ByVal sender As Object, ByVal e As KeyEventArgs)
        Public Event DragStatusChanged(ByVal DragOn As Boolean)
        Public Event DefaultAction_Changed(ByVal DefaultAction As Action)

        Private _CanEdit As Boolean = False
        Private Doc As Xml.XmlDocument
        Public Shared Content As New Panel
        Public ElementContext As ContextMenu
        Public ToolTip1 As New ToolTip
        Public Grid As Integer = 4

        Dim intCount As Integer = 0
        Public SelectedIndex As String = Nothing
        Dim blnDrag As Boolean = False
        Public DefaultAction As Action
        Dim MoveX As Integer
        Dim MoveY As Integer
        Public Elements As New Collection
        Public Containers As New Collection
        Dim Clipboard As Object
        Public UndoDelete As Panel
        Public UndoTop As Integer
        Public UndoLeft As Integer
        Public UndoWidth As Integer
        Public UndoHeight As Integer
        Public UndoIndex As String
        Dim blnLoaded As Boolean = False
        Public blnUndoable As Boolean = False

        Dim DivOutline As System.Drawing.Color = System.Drawing.Color.BlueViolet
        Dim DivBorderSize As Integer = 2

        Public ReadOnly Property SelectedElement() As Object
            Get
                If Not Me.SelectedIndex Is Nothing Then
                    Return Me.Elements(Me.SelectedIndex)
                Else
                    Return Nothing
                End If
            End Get
        End Property

        Public ReadOnly Property SelectedContainer() As Panel
            Get
                If Not Me.SelectedIndex Is Nothing Then
                    Return Me.Containers(Me.SelectedIndex)
                Else
                    Return Nothing
                End If
            End Get
        End Property

        Public ReadOnly Property CanEdit() As Boolean
            Get
                Return _CanEdit
            End Get
        End Property

        Public Enum Action
            Move
            Resize
            Text
        End Enum

        Public Sub New(ByRef Canvas As Panel, Optional ByVal CanEdit As Boolean = False)
            Me.Content = Canvas
            Me._CanEdit = CanEdit
            Me.SetDefaultAction(Action.Move)
            Me.blnLoaded = True
        End Sub

        Public Sub Clear()
            If Me.Containers.Count > 0
                Me.Elements = New Collection
                Me.Containers = New Collection
                Me.intCount = 0
                Me.Content.Controls.Clear()
            End If
        End Sub


        Public Sub SetDefaultAction(ByVal NewAction As Action)
            Dim Panel As Panel
            Select Case NewAction
                Case Action.Resize
                    DefaultAction = Action.Resize
                    ' Set label cursors
                    For Each Panel In Me.Containers
                        Panel.Cursor = System.Windows.Forms.Cursors.SizeNWSE
                    Next
                    ' Set all text boxes as readonly
                    Dim Element As Object
                    For Each Element In Me.Elements
                        Select Case Element.GetType.ToString
                            Case "System.Windows.Forms.TextBox"
                                Dim Text As TextBox = Element
                                Text.ReadOnly = True
                                Text.Cursor = Cursors.SizeNWSE
                            Case "CustomControls.TableView"
                                Dim TV As TableView = Element
                                TV.Cursor = Cursors.Default
                                'TV.ReadOnly = True
                        End Select
                    Next
                Case Action.Move
                    DefaultAction = Action.Move
                    ' Set label cursors
                    For Each Panel In Me.Containers
                        Panel.Cursor = System.Windows.Forms.Cursors.SizeAll
                    Next
                    ' Set all text boxes as readonly
                    Dim Element As Object
                    For Each Element In Me.Elements
                        Select Case Element.GetType.ToString
                            Case "System.Windows.Forms.TextBox"
                                Dim Text As TextBox = Element
                                Text.ReadOnly = True
                                Text.Cursor = Cursors.Arrow
                            Case "CustomControls.TableView"
                                Dim TV As TableView = Element
                                TV.Cursor = Cursors.Default
                                'TV.ReadOnly = True
                        End Select
                    Next
                Case Action.Text
                    DefaultAction = Action.Text
                    ' Set label cursors
                    For Each Panel In Me.Containers
                        Panel.Cursor = System.Windows.Forms.Cursors.IBeam
                    Next
                    ' Set all text boxes as editable
                    Dim Element As Object
                    For Each Element In Me.Elements
                        Select Case Element.GetType.ToString
                            Case "System.Windows.Forms.TextBox"
                                Dim Text As TextBox = Element
                                Text.ReadOnly = False
                                Text.Cursor = Cursors.IBeam
                            Case "CustomControls.TableView"
                                Dim TV As TableView = Element
                                TV.Cursor = Cursors.Default
                                'TV.ReadOnly = False
                        End Select
                    Next
            End Select
            RaiseEvent DefaultAction_Changed(NewAction)
        End Sub

        Private Function NextId() As Integer
            intCount += 1
            Return intCount
        End Function

        Public Sub SetValue(ByVal Variable As String, ByVal Value As String)
            Dim Element As Object
            For Each Element In Me.Elements
                If Element.GetType.ToString = "System.Windows.Forms.TextBox" Or Element.GetType.ToString = "CustomControls.TableView" Then
                    Element.Text = Element.Text.ToString.Replace(Variable, Value)
                End If
            Next
        End Sub

        Public Sub PopulateTable(ByVal DataSource As String, ByVal Data As DataTable)
            Dim Index As String = Me.FindTableWithDataSource(DataSource)
            If Not Index Is Nothing Then
                Dim Table As TableView = Me.Elements(Index)
                Dim Cols As TableView.ColumnsCollection = Table.Columns
                Dim Row As DataRow
                For Each Row In Data.Rows
                    Dim Fields As New Collection
                    For i As Integer = 1 To Cols.Items.Count
                        Try
                            Fields.Add(Row.Item(CType(Cols.Items(i), TableView.Column).MappingName))
                        Catch ex As Exception
                            MsgBox("Column does not exist, template is probably wrong. " & ex.ToString)
                        End Try
                    Next
                    Table.Rows.Add(Fields)
                Next
            End If
        End Sub

        Public Function FindTableWithDataSource(ByVal DataSource As String) As String
            Dim El As Object
            For Each El In Me.Elements
                If El.GetType.ToString = "CustomControls.TableView" Then
                    If CType(El, TableView).DataSource = DataSource Then
                        Return El.Tag
                    End If
                End If
            Next
            Return Nothing
        End Function

        Public Function GetElementsByType(ByVal Type As String) As String()
            Dim El As Object
            Dim strReturn As String() = Nothing
            Dim i As Integer = 0
            For Each El In Me.Elements
                Dim ThisType As String = El.GetType.ToString
                If ThisType = Type Then
                    Try
                        ReDim Preserve strReturn(strReturn.Length)
                    Catch ex As NullReferenceException
                        ReDim strReturn(0)
                    Catch
                        Return Nothing
                    End Try
                    strReturn(i) = El.Tag
                    i += 1
                End If
            Next
            Return strReturn
        End Function

        Public Function GetFirstElementByType(ByVal Type As String) As String
            Dim El As String() = Me.GetElementsByType(Type)
            Dim ReturnVal As String
            Try
                ReturnVal = El(0)
            Catch
                ReturnVal = Nothing
            End Try
            Return ReturnVal
        End Function

        Public Function GetElementById(ByVal ID As String) As Object
            For i As Integer = 1 To Me.Elements.Count
                If Me.Elements.Item(i).Name = ID Then
                    Return Me.Elements.Item(i).Tag
                End If
            Next
            Return Nothing
        End Function

        Public Function NameExists(ByVal Name As String) As Boolean
            For i As Integer = 1 To Me.Elements.Count
                If Me.Elements.Item(i).Name = Name Then
                    Return True
                End If
            Next
            Return False
        End Function


#Region "Load XML"

        Public Sub LoadXml(ByVal xd As Xml.XmlDocument)
            ' Set XMl Doc
            Me.Doc = xd
            ' Define variables
            Dim Nodes As XmlNodeList = Doc.SelectNodes("/html/body/page/div")
            Dim Node As XmlNode
            Dim Order(Nodes.Count) As String
            ' Loop through xml
            For Each Node In Nodes
                Select Case Node.SelectSingleNode("@class").InnerText
                    Case "Label"
                        Try
                            BuildLabel(Node.SelectSingleNode("@style").InnerText, Node.InnerText, Node.SelectSingleNode("@id").InnerText)
                        Catch ex As Exception
                            Throw New Exception("Error buliding label. " & ex.ToString)
                        End Try
                    Case "Box"
                        Try
                            BuildBox(Node.SelectSingleNode("@style").InnerText, Node.SelectSingleNode("@id").InnerText)
                        Catch ex As Exception
                            Throw New Exception("Error buliding box. " & ex.ToString)
                        End Try
                    Case "Image"
                        Try
                            BuildImage(Node.SelectSingleNode("@style").InnerText, Node.SelectSingleNode("img/@src").InnerText, Node.SelectSingleNode("@id").InnerText)
                        Catch ex As Exception
                            Throw New Exception("Error buliding image. " & ex.ToString)
                        End Try
                    Case "Signature"
                        Try
                            BuildSignatureBox(Node.SelectSingleNode("@style").InnerText, Node.SelectSingleNode("img/@src").InnerText, Node.SelectSingleNode("@id").InnerText)
                        Catch ex As Exception
                            Throw New Exception("Error buliding signature. " & ex.ToString)
                        End Try
                    Case "Table"
                        'Dim Caption As XmlElement = Node.SelectSingleNode("table/caption")
                        Dim Table As TableView
                        Try
                            Table = BuildTable(Node.SelectSingleNode("@style").InnerText, Node.SelectSingleNode("@id").InnerText)
                        Catch ex As Exception
                            Throw New Exception("Error buliding table. " & ex.ToString)
                        End Try
                        Dim TableStyle As Hashtable
                        Try
                            TableStyle = Me.ExtractStyle(Node.SelectSingleNode("@style").InnerText)
                        Catch
                            Throw New Exception("Error extracting table style.")
                        End Try
                        ' Columns and Data
                        Dim Columns As XmlNodeList = Node.FirstChild.FirstChild.SelectNodes("th")
                        Dim Rows As XmlNodeList = Node.SelectNodes("table/tr")
                        Dim Col As XmlNode
                        For Each Col In Columns
                            Dim Map As String = ""
                            Try
                                Map = Col.SelectSingleNode("@map").InnerText
                            Catch ex As Exception
                                ' Nothing
                            End Try
                            Dim Format As String = ""
                            Try
                                Format = Col.SelectSingleNode("@format").InnerText
                            Catch ex As Exception
                                ' Nothing
                            End Try
                            Try
                                Table.Columns.Add(Col.InnerText, Col.SelectSingleNode("@width").InnerText, Map, Format)
                            Catch
                                Throw New Exception("Error adding column.")
                            End Try
                        Next
                        Try
                            Table.DataSource = Node.SelectSingleNode("table/@src").InnerText
                        Catch
                            Throw New Exception("Error setting source.")
                        End Try
                        Try
                            Table.RowLimit = Node.SelectSingleNode("table/@rowlimit").InnerText
                        Catch
                            Table.RowLimit = 0
                        End Try
                        Try
                            Table.CellSpacing = Node.SelectSingleNode("table/@cellspacing").InnerText
                        Catch
                            Table.CellSpacing = 1
                        End Try
                        Try
                            Table.CellPadding = Node.SelectSingleNode("table/@cellpadding").InnerText
                        Catch
                            Table.CellPadding = 1
                        End Try
                        ' Table Cell Font
                        Try
                            Table.CellFont = TableStyle.Item("font-family").ToString
                        Catch ex As Exception
                            Table.CellFont = "Arial"
                        End Try
                        Try
                            Table.CellFontSize = TableStyle.Item("font-size").ToString
                        Catch ex As Exception
                            Table.CellFontSize = 8
                        End Try
                        ' Table Border
                        Try
                            If TableStyle.Item("border-style") = "solid" Then
                                Table.BorderStyle = BorderStyle.FixedSingle
                            ElseIf TableStyle.Item("border-style") = "ridge" Then
                                Table.BorderStyle = BorderStyle.Fixed3D
                            Else
                                Table.BorderStyle = BorderStyle.None
                            End If
                        Catch
                            Table.BorderStyle = BorderStyle.None
                        End Try
                        ' Rows
                        If Rows.Count > 1 Then
                            For i As Integer = 1 To Rows.Count - 1
                                Dim Fields As New Collection
                                Dim Cells As XmlNodeList = Rows(i).SelectNodes("td")
                                For Each Cell As XmlNode In Cells
                                    Fields.Add(Cell.InnerText)
                                Next
                                Table.Rows.Add(Fields)
                            Next
                        End If
                        ' Caption
                        Table.CaptionVisble = False
                        'If Not Caption Is Nothing Then
                        '    Dim CaptionStyle As Hashtable = Me.ExtractStyle(Node.SelectSingleNode("table/caption/@style").InnerText)
                        '    Table.CaptionText = Node.SelectSingleNode("table/caption").InnerText.Replace(ControlChars.Lf, ControlChars.CrLf)

                        '    Select Case CaptionStyle.Item("text-align").ToString
                        '        Case "center"
                        '            Table.CaptionTextAlign = HorizontalAlignment.Center
                        '        Case "left"
                        '            Table.CaptionTextAlign = HorizontalAlignment.Left
                        '        Case Else
                        '            Table.CaptionTextAlign = HorizontalAlignment.Right
                        '    End Select
                        '    Table.CaptionBackColor = Color.FromArgb(CaptionStyle.Item("background-color"))
                        '    Table.CaptionForeColor = Color.FromArgb(CaptionStyle.Item("color"))
                        '    Table.CaptionPadding = CaptionStyle.Item("padding")
                        '    Dim FontStyle As FontStyle
                        '    ' Bold
                        '    If CaptionStyle.Item("font-weight") = "bold" Then
                        '        FontStyle = FontStyle.Bold
                        '    End If
                        '    ' Italic
                        '    If CaptionStyle.Item("font-style") = "italic" Then
                        '        FontStyle = FontStyle Or FontStyle.Italic
                        '    End If
                        '    ' Underline
                        '    If CaptionStyle.Item("text-decoration") = "underline" Then
                        '        FontStyle = FontStyle Or FontStyle.Underline
                        '    End If
                        '    ' Set font
                        '    Table.CaptionFont = New Font(New FontFamily(CaptionStyle.Item("font-family").ToString), CaptionStyle.Item("font-size"), FontStyle)
                        '    If CaptionStyle.Item("caption-side") = "bottom" Then
                        '        Table.CaptionSide = CustomControls.TableView.CaptionEnum.Bottom
                        '    Else
                        '        Table.CaptionSide = CustomControls.TableView.CaptionEnum.Top
                        '    End If
                        'Else
                        '    Table.CaptionVisble = False
                        'End If
                End Select
                Try
                    Order(CType(ExtractSingleStyle(Node, "z-index"), Integer)) = "Element" & Me.intCount.ToString
                Catch

                End Try
            Next
            ' Reset order
            For i As Integer = 0 To Order.Length - 1
                Try
                    If Not Order(i) = Nothing Then
                        CType(Me.Containers.Item(Order(i)), Panel).BringToFront()
                    End If
                Catch ex As Exception
                    Dim ErrorBox As New MyCore.Gravity.ErrorBox("Error resorting at index " & i, "Z-Index Sorting Error", ex.ToString, 1)
                End Try
            Next
        End Sub

        Private Function ExtractStyle(ByVal Style As String) As Hashtable
            Dim Table As New Hashtable
            Dim SplitStyle As String() = Style.Split(";")
            Dim Attrib(1) As String
            Dim StyleAttrib As String
            For Each StyleAttrib In SplitStyle
                If StyleAttrib.Trim.Length > 0 Then
                    Attrib = StyleAttrib.Split(":")
                    Select Case Attrib(0).Trim
                        Case "font-size", "border-width", "left", "top", "padding", "margin", "width", "height"
                            Attrib(1) = Attrib(1).Replace("pt", "")
                            Attrib(1) = Attrib(1).Replace("em", "")
                            Attrib(1) = Attrib(1).Replace("px", "")
                            Attrib(1) = Attrib(1).Replace("cm", "")
                            Attrib(1) = Attrib(1).Replace("ex", "")
                            Attrib(1) = Attrib(1).Replace("mm", "")
                            Attrib(1) = Attrib(1).Replace("in", "")
                        Case "color", "background-color", "border-color"
                            Dim SplitColor(2) As String
                            Attrib(1) = Attrib(1).Replace("rgb(", "")
                            Attrib(1) = Attrib(1).Replace(")", "")
                            SplitColor = Attrib(1).Split(",")
                            Attrib(1) = Color.FromArgb(SplitColor(0).Trim, SplitColor(1).Trim, SplitColor(2).Trim).ToArgb.ToString
                    End Select
                    Table.Add(Attrib(0).Trim, Attrib(1).Trim)
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

#End Region

#Region "Construction of Elements"

        Public Sub BuildImage(ByVal strStyle As String, Optional ByVal Src As String = Nothing, Optional ByVal ID As String = Nothing)
            Dim Style As Hashtable = ExtractStyle(strStyle)
            Dim Image As New PictureBox
            Dim Index As String = Me.BuildElement(Image, Style, ID)
            ' Set up picture box
            Image.BorderStyle = BorderStyle.None
            Image.Cursor = Windows.Forms.Cursors.Hand
            Image.BackColor = System.Drawing.Color.White
            If Not Src = Nothing Then
                If Src.StartsWith("data:image/gif;base64,") Then
                    Dim strBytes As String = Src.Substring(22)
                    Dim Bytes As Byte() = System.Convert.FromBase64String(strBytes)
                    Dim MS As New System.IO.MemoryStream(Bytes)
                    Image.Image = Image.Image.FromStream(MS)
                End If
            End If
            If Style.Item("border-style") = "solid" Then
                Image.BorderStyle = BorderStyle.FixedSingle
            Else
                Image.BorderStyle = BorderStyle.None
            End If
            AddHandler Image.DoubleClick, AddressOf Element_DoubleClick
            AddHandler Image.Click, AddressOf Label_Enter
            AddHandler Image.MouseDown, AddressOf Element_MouseDown
            AddHandler Image.MouseUp, AddressOf Me.Element_MouseUp
            AddHandler Image.MouseMove, AddressOf Element_MouseMove
        End Sub

        Public Sub BuildBox(ByVal strStyle As String, Optional ByVal ID As String = Nothing)
            Dim Style As Hashtable = ExtractStyle(strStyle)
            Dim Label As New Label
            Dim Index As String = Me.BuildElement(Label, Style, ID)
            ' Set up picture box
            Label.BorderStyle = BorderStyle.None
            Try
                Label.BackColor = System.Drawing.Color.FromArgb(Style.Item("background-color").ToString)
            Catch
                Label.BackColor = Color.Black
            End Try
            Label.Text = ""
            Label.AutoSize = False
            If Style.Item("border-style") = "solid" Then
                Label.BorderStyle = BorderStyle.FixedSingle
            Else
                Label.BorderStyle = BorderStyle.None
            End If
            AddHandler Label.DoubleClick, AddressOf Element_DoubleClick
            AddHandler Label.Click, AddressOf Label_Enter
            AddHandler Label.MouseDown, AddressOf Element_MouseDown
            AddHandler Label.MouseUp, AddressOf Me.Element_MouseUp
            AddHandler Label.MouseMove, AddressOf Element_MouseMove
            AddHandler Label.KeyDown, AddressOf Element_KeyDown
        End Sub

        Public Function BuildTable(ByVal strStyle As String, Optional ByVal ID As String = Nothing) As TableView
            Dim Style As Hashtable = ExtractStyle(strStyle)
            Dim Table As New TableView
            Dim Index As String = Me.BuildElement(Table, Style, ID)
            Dim FontStyle As New FontStyle
            ' Set up Label
            Table.BorderStyle = BorderStyle.None
            Table.Cursor = Windows.Forms.Cursors.Default
            Table.FullRowSelect = True
            Table.View = View.Details
            Table.HeaderStyle = ColumnHeaderStyle.Nonclickable
            If Not Me.CanEdit Then
                Table.CaptionLocked = True
            End If
            ' The border is no longer to be controlled in the DIV, but in the TABLE tag
            'If Style.Item("border-style") = "solid" Then
            'Table.BorderStyle = BorderStyle.FixedSingle
            'Else
            '    Table.BorderStyle = BorderStyle.None
            'End If
            AddHandler Table.DoubleClick, AddressOf Element_DoubleClick
            AddHandler Table.Enter, AddressOf Label_Enter
            AddHandler Table.MouseDown, AddressOf Element_MouseDown
            AddHandler Table.MouseUp, AddressOf Me.Element_MouseUp
            AddHandler Table.MouseMove, AddressOf Element_MouseMove
            ' Return
            Return Table
        End Function

        Public Sub BuildSignatureBox(ByVal strStyle As String, Optional ByVal ImageSrc As String = Nothing, Optional ByVal ID As String = Nothing)
            Dim Style As Hashtable = ExtractStyle(strStyle)
            Dim Signature As New SignatureBox
            Dim Index As String = Me.BuildElement(Signature, Style, ID)
            Dim FontStyle As New FontStyle
            ' Format signature box
            Signature.Cursor = Windows.Forms.Cursors.Cross
            Signature.BackColor = Me.Content.BackColor
            If Not ImageSrc Is Nothing Then
                If ImageSrc.StartsWith("data:image/gif;base64,") Then
                    Dim strBytes As String = ImageSrc.Substring(22)
                    Signature.RawData = strBytes
                Else
                    Signature.FileName = ImageSrc
                End If
            End If
            AddHandler Signature.DoubleClick, AddressOf Element_DoubleClick
            AddHandler Signature.Enter, AddressOf Label_Enter
            AddHandler Signature.MouseDown, AddressOf Element_MouseDown
            AddHandler Signature.MouseUp, AddressOf Me.Element_MouseUp
            AddHandler Signature.MouseMove, AddressOf Element_MouseMove
        End Sub

        Public Sub BuildLabel(ByVal strStyle As String, Optional ByVal Content As String = "", Optional ByVal ID As String = Nothing)
            Dim Style As Hashtable = ExtractStyle(strStyle)
            Dim Label As New TextBox
            Dim Index As String = Me.BuildElement(Label, Style, ID)
            Dim FontStyle As New FontStyle
            ' Format label
            If Not Me.CanEdit Then
                Label.ReadOnly = True
                Label.Cursor = System.Windows.Forms.Cursors.Arrow
                AddHandler Label.Enter, AddressOf Me.TextBox_Enter
            End If
            Label.Multiline = True
            Label.BorderStyle = BorderStyle.None
            Label.BackColor = System.Drawing.Color.FromArgb(Style.Item("background-color").ToString)
            Label.ForeColor = System.Drawing.Color.FromArgb(Style.Item("color").ToString)
            Label.Text = Content.Replace(ControlChars.Lf, ControlChars.CrLf)
            If Style.Item("border-style") = "solid" Then
                Label.BorderStyle = BorderStyle.FixedSingle
            Else
                Label.BorderStyle = BorderStyle.None
            End If
            ' Text Align
            Select Case Style.Item("text-align").ToString.ToLower
                Case "left"
                    Label.TextAlign = HorizontalAlignment.Left
                Case "right"
                    Label.TextAlign = HorizontalAlignment.Right
                Case Else
                    Label.TextAlign = HorizontalAlignment.Center
            End Select
            ' Bold
            If Style.Item("font-weight") = "bold" Then
                FontStyle = FontStyle.Bold
            End If
            ' Italic
            If Style.Item("font-style") = "italic" Then
                FontStyle = FontStyle Or FontStyle.Italic
            End If
            ' Underline
            If Style.Item("text-decoration") = "underline" Then
                FontStyle = FontStyle Or FontStyle.Underline
            End If
            ' Set font
            Try
                Label.Font = New Font(New FontFamily(Style.Item("font-family").ToString), Style.Item("font-size"), FontStyle)
            Catch
                Label.Font = New Font(New FontFamily(System.Drawing.Text.GenericFontFamilies.SansSerif), Style.Item("font-size"), FontStyle)
            End Try
            ' Event Handlers
            AddHandler Label.DoubleClick, AddressOf Element_DoubleClick
            AddHandler Label.Enter, AddressOf Label_Enter
            AddHandler Label.KeyDown, AddressOf Element_KeyDown
            AddHandler Label.MouseDown, AddressOf Element_MouseDown
            AddHandler Label.MouseUp, AddressOf Me.Element_MouseUp
            AddHandler Label.MouseMove, AddressOf Element_MouseMove
        End Sub

        Private Function BuildElement(ByVal Element As Object, ByVal Style As Hashtable, Optional ByVal ID As String = Nothing) As String
            Dim Panel As New Panel
            Dim Index As String = "Element" & NextId()
            Dim Name As String = ""
            ' Name Exists?
            If ID = Nothing Then
                ID = Index
                Do Until Not Me.NameExists(ID)
                    ID &= Index
                Loop
            ElseIf Me.NameExists(ID) Then
                ID &= Index
                Do Until Not Me.NameExists(ID)
                    ID &= Index
                Loop
            End If
            ' Set up Panel
            Panel.Name = ID & "Panel"
            Panel.Tag = Index
            If Me._CanEdit Then
                Panel.Left = Style.Item("left") - DivBorderSize
                Panel.Top = Style.Item("top") - DivBorderSize
                Panel.Width = Style.Item("width") + DivBorderSize * 2
                Panel.Height = Style.Item("height") + DivBorderSize * 2
                Panel.ContextMenu = Me.ElementContext
                Panel.BackColor = DivOutline
                Panel.DockPadding.All = 2
            Else
                Panel.Left = Style.Item("left")
                Panel.Top = Style.Item("top")
                Panel.Width = Style.Item("width")
                Panel.Height = Style.Item("height")
            End If
            ' Set up Element
            Element.Name = ID
            Element.Dock = DockStyle.Fill
            Element.Tag = Index
            ' Tooltip
            If Me._CanEdit Then
                Me.ToolTip1.SetToolTip(Element, Element.Name)
                Me.ToolTip1.SetToolTip(Panel, Element.Name)
            End If
            ' Cursor
            If Me.CanEdit Then
                If DefaultAction = Action.Move Then
                    Panel.Cursor = System.Windows.Forms.Cursors.SizeAll
                ElseIf DefaultAction = Action.Resize Then
                    Panel.Cursor = System.Windows.Forms.Cursors.SizeNWSE
                End If
            End If
            ' Context menu
            Element.ContextMenu = Me.ElementContext
            Panel.ContextMenu = Me.ElementContext
            ' Add Label to Panel
            Panel.Controls.Add(Element)
            ' Add Panel to Form
            Me.Content.Controls.Add(Panel)
            ' Add to collections
            Me.Elements.Add(Element, Index)
            Me.Containers.Add(Panel, Index)
            ' Bring to font
            Panel.BringToFront()
            ' Add Event handlers
            AddHandler Panel.Enter, AddressOf Panel_Enter
            AddHandler Panel.MouseDown, AddressOf Element_MouseDown
            AddHandler Panel.MouseUp, AddressOf Me.Element_MouseUp
            AddHandler Panel.MouseMove, AddressOf Element_MouseMove
            Return Index
        End Function

#End Region

#Region "Export XML"

        Public Function BMPToBytes(ByVal bmp As Image) As Byte()
            Dim ms As New System.IO.MemoryStream
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Gif)
            Dim abyt(ms.Length - 1) As Byte
            ms.Seek(0, IO.SeekOrigin.Begin)
            ms.Read(abyt, 0, ms.Length)
            Return abyt
        End Function

        Public Function Export(Optional ByVal Restraints As TableRestraint() = Nothing) As XmlDocument

            Dim Body As XmlNode = Me.Doc.SelectSingleNode("/html/body")
            Body.RemoveAll()

            Dim Element As Object
            For Each Element In Me.Elements
                If Element.GetType.ToString = "System.Windows.Forms.TextBox" Then
                    Dim Div As XmlElement = Me.Doc.CreateElement("div")
                    Div.Attributes.Append(NewAttribute("id", Element.Name))
                    Div.Attributes.Append(NewAttribute("class", "Label"))
                    Div.Attributes.Append(NewAttribute("style", BuildStyleString(Element)))
                    Dim SplitContent As String() = CType(Element, TextBox).Text.Split(ControlChars.CrLf)
                    Dim Line As String
                    For Each Line In SplitContent
                        Div.AppendChild(Me.Doc.CreateTextNode(Line))
                        Div.AppendChild(Me.Doc.CreateElement("br"))
                    Next
                    Body.AppendChild(Div)
                ElseIf Element.GetType.ToString = "System.Windows.Forms.Label" Then
                    Dim Div As XmlElement = Me.Doc.CreateElement("div")
                    Div.Attributes.Append(NewAttribute("id", Element.Name))
                    Div.Attributes.Append(NewAttribute("class", "Box"))
                    Div.Attributes.Append(NewAttribute("style", BuildStyleString(Element)))
                    Body.AppendChild(Div)
                ElseIf Element.GetType.ToString = "System.Windows.Forms.PictureBox" Then
                    Dim Div As XmlElement = Me.Doc.CreateElement("div")
                    Dim Img As XmlElement = Me.Doc.CreateElement("img")
                    Dim Box As Windows.Forms.PictureBox = Element
                    Div.Attributes.Append(NewAttribute("style", BuildStyleString(Element)))
                    Div.Attributes.Append(NewAttribute("class", "Image"))
                    Div.Attributes.Append(NewAttribute("id", Element.Name))
                    Dim Bytes As Byte()
                    Try
                        Bytes = Me.BMPToBytes(Box.Image)
                    Catch
                        Dim bmp As New System.Drawing.Bitmap(Box.Parent.Width, Box.Parent.Height)
                        Dim gfx As Graphics
                        gfx = gfx.FromImage(bmp)
                        gfx.Clear(Color.White)
                        Bytes = Me.BMPToBytes(bmp)
                    End Try
                    Dim Raw As String = System.Convert.ToBase64String(Bytes)
                    Img.Attributes.Append(NewAttribute("src", "data:image/gif;base64," & Raw))
                    Img.Attributes.Append(NewAttribute("style", "width: 100%; height: 100%"))
                    Div.AppendChild(Img)
                    Body.AppendChild(Div)
                ElseIf Element.GetType.ToString = "CustomControls.TableView" Then
                        Dim Div As XmlElement = Me.Doc.CreateElement("div")
                        Dim Table As XmlElement = Me.Doc.CreateElement("table")
                        Dim Caption As XmlElement = Me.Doc.CreateElement("caption")
                        Dim CaptionStyle As String
                    Dim tv As TableView = CType(Element, TableView)
                        Div.Attributes.Append(NewAttribute("style", BuildStyleString(Element, True)))
                        Div.Attributes.Append(NewAttribute("id", Element.Name))
                        Div.Attributes.Append(NewAttribute("class", "Table"))
                        Dim strStyle As String = "width: 100%; "
                        If tv.BorderStyle = BorderStyle.FixedSingle Then
                            strStyle &= "border-style: solid; border-width: 2px; border-color: #000000; "
                        ElseIf tv.BorderStyle = BorderStyle.Fixed3D Then
                            strStyle &= "border-style: ridge; border-width: 4px; "
                        Else
                            strStyle &= "border-style: none; "
                        End If
                        strStyle &= "font-size: " & tv.CellFontSize & "px; "
                        strStyle &= "font-family: " & tv.CellFont & "; "
                        Table.Attributes.Append(NewAttribute("style", strStyle))
                        Table.Attributes.Append(NewAttribute("src", Element.DataSource))
                        Table.Attributes.Append(NewAttribute("rowlimit", tv.RowLimit))
                        Table.Attributes.Append(NewAttribute("cellpadding", tv.CellPadding))
                        Table.Attributes.Append(NewAttribute("cellspacing", tv.CellSpacing))
                        ' Column header row
                        Dim Tr As XmlElement = Me.Doc.CreateElement("tr")
                        Tr.Attributes.Append(NewAttribute("valign", "top"))
                        ' Columns
                    Dim Col As TableView.Column
                        For Each Col In tv.Columns.Items
                            Dim Th As XmlElement = Me.Doc.CreateElement("th")
                            Th.InnerText = Col.Header.Text
                            Th.Attributes.Append(NewAttribute("map", Col.MappingName))
                            If Col.Format.Length > 0 Then
                                Th.Attributes.Append(NewAttribute("format", Col.Format))
                            End If
                            Th.Attributes.Append(NewAttribute("width", Col.Header.Width))
                            Th.Attributes.Append(NewAttribute("align", "left"))
                            Tr.AppendChild(Th)
                        Next
                        Table.AppendChild(Tr)
                        'Rows
                    Dim Row As TableView.Row
                        Dim Count As Integer = 0
                        For Each Row In tv.Rows.Items
                            ' If there is a table restraint (limit) on it, determine whether 
                            ' to show or hide this row
                            Dim showit As Boolean = False
                            Dim r As TableRestraint = Me.IsRestraint(Restraints, tv.DataSource)
                            If r Is Nothing Then
                                showit = True
                            ElseIf Count >= r.Start And Count - r.Start < r.Limit Then
                                showit = True
                            End If
                            ' If we're showing it
                            If showit Then
                                Tr = Me.Doc.CreateElement("tr")
                                Tr.Attributes.Append(NewAttribute("valign", "top"))
                                For i As Integer = 1 To Row.Fields.Count
                                    Dim strFormat As String = ""
                                    Try
                                        Col = tv.Columns.Items(i)
                                        strFormat = Col.Format
                                    Catch
                                    End Try
                                    Dim Td As XmlElement = Me.Doc.CreateElement("td")
                                    Td.InnerText = Format(Row.Fields(i), strFormat)
                                    Tr.AppendChild(Td)
                                Next
                                Table.AppendChild(Tr)
                            End If
                            ' Increment counter
                            Count += 1
                        Next
                        ' Caption
                        If tv.CaptionVisble Then
                        If tv.CaptionSide = TableView.CaptionEnum.Bottom Then
                            CaptionStyle = "caption-side: bottom; "
                        Else
                            CaptionStyle = "caption-side: top; "
                        End If
                            If tv.CaptionTextAlign = HorizontalAlignment.Center Then
                                CaptionStyle &= "text-align: center; "
                            ElseIf tv.CaptionTextAlign = HorizontalAlignment.Left Then
                                CaptionStyle &= "text-align: left; "
                            Else
                                CaptionStyle &= "text-align: right; "
                            End If
                            If tv.BorderStyle = BorderStyle.FixedSingle Then
                                CaptionStyle &= "border-left: solid 2px rgb(0, 0, 0); "
                                CaptionStyle &= "border-right: solid 2px rgb(0, 0, 0); "
                                CaptionStyle &= "border-bottom: solid 2px rgb(0, 0, 0); "
                                CaptionStyle &= "border-top: none; "
                            Else
                                CaptionStyle &= "border: none; "
                            End If
                            CaptionStyle &= " font-size: " & tv.CaptionFont.Size & "pt; "
                            CaptionStyle &= " color: rgb(" & tv.CaptionForeColor.R & ", " & tv.CaptionForeColor.G & ", " & tv.CaptionForeColor.B & "); "
                            CaptionStyle &= " background-color: rgb(" & tv.CaptionBackColor.R & ", " & tv.CaptionBackColor.G & ", " & tv.CaptionBackColor.B & "); "
                            If tv.CaptionFont.Bold Then
                                CaptionStyle &= "font-weight: bold; "
                            Else
                                CaptionStyle &= "font-weight: normal; "
                            End If
                            If tv.CaptionFont.Italic Then
                                CaptionStyle &= "font-style: italic; "
                            Else
                                CaptionStyle &= "font-style: normal; "
                            End If
                            If tv.CaptionFont.Underline Then
                                CaptionStyle &= "text-decoration: underline; "
                            Else
                                CaptionStyle &= "text-decoration: none; "
                            End If
                            CaptionStyle &= "padding: " & tv.CaptionPadding & "; "
                            CaptionStyle &= "font-family: " & tv.CaptionFont.Name & "; "
                            ' Could put this in but really it's unecessary because caption
                            ' will be only as big as its content
                            'CaptionStyle &= "height: " & tv.CaptionSize & "; "
                            Caption.Attributes.Append(NewAttribute("style", CaptionStyle))
                            ' Internet Explorer Fix: IE6 doesn't support the caption-side CSS property
                        If tv.CaptionSide = TableView.CaptionEnum.Bottom Then
                            Caption.Attributes.Append(NewAttribute("valign", "bottom"))
                        Else
                            Caption.Attributes.Append(NewAttribute("valign", "top"))
                        End If
                        Dim SplitContent As String() = CType(Element, TableView).CaptionText.Split(ControlChars.CrLf)
                            Dim Line As String
                            For Each Line In SplitContent
                                Caption.AppendChild(Me.Doc.CreateTextNode(Line))
                                Caption.AppendChild(Me.Doc.CreateElement("br"))
                            Next
                        End If
                        If tv.CaptionVisble Then
                            Table.AppendChild(Caption)
                        End If
                        Div.AppendChild(Table)
                        Body.AppendChild(Div)
                ElseIf Element.GetType.ToString = "CustomControls.SignatureBox" Then
                        Dim Div As XmlElement = Me.Doc.CreateElement("div")
                        Dim Img As XmlElement = Me.Doc.CreateElement("img")
                        Div.Attributes.Append(NewAttribute("style", BuildStyleString(Element)))
                        Div.Attributes.Append(NewAttribute("class", "Signature"))
                        Div.Attributes.Append(NewAttribute("id", Element.Name))
                        Dim Raw As String = Element.RawData
                        Img.Attributes.Append(NewAttribute("src", "data:image/gif;base64," & Raw))
                        Img.Attributes.Append(NewAttribute("style", "width: 100%; height: 100%"))
                        Div.AppendChild(Img)
                        Body.AppendChild(Div)
                End If
            Next

            Me.Doc.PreserveWhitespace = True

            Return Me.Doc

        End Function

        Public Overrides Function ToString() As String
            Dim Doc As XmlDocument = Me.Export
            Return Doc.SelectSingleNode("/").OuterXml
        End Function

        Private Function NewAttribute(ByVal Name As String, ByVal Value As String) As XmlAttribute
            Dim Attrib As XmlAttribute = Me.Doc.CreateAttribute(Name)
            Attrib.InnerText = Value
            Return Attrib
        End Function

        Private Function BuildStyleString(ByVal Element As Object, Optional ByVal IgnoreBorder As Boolean = False) As String
            Dim Style As String
            ' Location and size
            Style = "position: absolute; " & _
                "left: " & Element.Parent.Left + DivBorderSize & "; " & _
                "top: " & Element.Parent.Top + DivBorderSize & "; " & _
                "width: " & Element.Parent.Width - DivBorderSize * 2 & "; " & _
                "height: " & Element.Parent.Height - DivBorderSize * 2 & "; " & _
                "z-index: " & CType(Element.Parent.Parent, Panel).Controls.Count - CType(Element.Parent.Parent, Panel).Controls.GetChildIndex(Element.Parent) & "; "
            ' Border
            If Not IgnoreBorder Then
                If Element.BorderStyle = BorderStyle.FixedSingle Then
                    Style &= " border-style: solid; "
                    Style &= " border-width: 2px; "
                    Style &= " border-color: rgb(0, 0, 0); "
                Else
                    Style &= " border-style: none; "
                End If
            End If
            If Element.GetType.ToString = "System.Windows.Forms.TextBox" Then
                Style &= "font-family: " & Element.Font.Name & "; " & _
                    "font-size: " & Element.Font.Size & "pt; " & _
                    "color: rgb(" & Element.ForeColor.R & ", " & Element.ForeColor.G & ", " & Element.ForeColor.B & "); " & _
                    "background-color: rgb(" & Element.BackColor.R & ", " & Element.BackColor.G & ", " & Element.BackColor.B & "); "
                ' Text Alignment
                If Element.TextAlign = HorizontalAlignment.Left Then
                    Style &= " text-align: left; "
                ElseIf Element.TextAlign = HorizontalAlignment.Right Then
                    Style &= " text-align: right; "
                Else
                    Style &= " text-align: center; "
                End If
                ' Underline
                If Element.Font.Underline Then
                    Style &= " text-decoration: underline; "
                Else
                    Style &= " text-decoration: none; "
                End If
                ' Italic
                If Element.Font.Italic Then
                    Style &= " font-style: italic; "
                Else
                    Style &= " font-style: normal; "
                End If
                ' Bold
                If Element.Font.Bold Then
                    Style &= " font-weight: bold; "
                Else
                    Style &= " font-weight: normal; "
                End If
            ElseIf Element.GetType.ToString = "CustomControls.TableView" Then
                Style &= " overflow: hidden; "
            ElseIf Element.GetType.ToString = "System.Windows.Forms.Label" Then
                Style &= " background-color: rgb(" & Element.BackColor.R & ", " & Element.BackColor.G & ", " & Element.BackColor.B & "); "
                Style &= " overflow: hidden; "
            End If
            Return Style
        End Function

        Private Function IsRestraint(ByVal Restraints As TableRestraint(), ByVal Name As String) As TableRestraint
            If Not Restraints Is Nothing Then
                Dim r As TableRestraint
                For Each r In Restraints
                    If r.Src = Name Then
                        Return r
                    End If
                Next
            End If
            Return Nothing
        End Function

#End Region

#Region "Copy/Cut/Paste"

        Private Sub AddToClipboard(ByVal obj As Object)
            Me.Clipboard = obj
        End Sub

        Public Sub Copy()
            If Not Me.SelectedIndex = Nothing Then
                Me.AddToClipboard(Me.Containers.Item(Me.SelectedIndex))
            End If
        End Sub

        Public Sub Paste(Optional ByVal obj As Panel = Nothing)
            Dim Panel As Panel
            If obj Is Nothing Then
                Panel = CType(Me.Clipboard, Panel)
            Else
                Panel = obj
            End If
            Try
                If Not Panel Is Nothing Then
                    Select Case Panel.Controls(0).GetType.ToString
                        Case "System.Windows.Forms.TextBox"
                            Me.BuildLabel(Me.BuildStyleString(Panel.Controls(0)), Panel.Controls(0).Text)
                        Case "System.Windows.Forms.Label"
                            Me.BuildBox(Me.BuildStyleString(Panel.Controls(0)))
                        Case "System.Windows.Forms.PictureBox"
                            Me.BuildImage(Me.BuildStyleString(Panel.Controls(0)), Panel.Controls(0).Text)
                        Case "System.Windows.Forms.ListView"
                            Dim Table As TableView = Me.BuildTable(Me.BuildStyleString(Panel.Controls(0)))
                            Dim Col As ColumnHeader
                            For Each Col In CType(Panel.Controls(0), TableView).Columns.Items
                                Table.Columns.Add(Col.Text, Col.Width, "")
                            Next
                        Case "CustomControls.SignatureBox"
                            Me.BuildSignatureBox(Me.BuildStyleString(Panel.Controls(0)), Panel.Controls(0).Text)
                    End Select
                End If
            Catch
            End Try
        End Sub

        Public Sub Delete(Optional ByVal Die As String = Nothing)
            If Die = Nothing Then
                If Not Me.SelectedIndex = Nothing Then
                    Die = Me.SelectedIndex
                End If
            End If
            If Not Die = Nothing Then
                Dim MarkedForDeath As String = Die
                If Not Me.Containers.Item(MarkedForDeath) Is Nothing Then
                    Me.UndoDelete = Me.Containers.Item(MarkedForDeath)
                    Me.blnUndoable = True
                    Me.UnselectElement()
                    Me.Content.Controls.Remove(Me.Containers.Item(MarkedForDeath))
                    Me.Containers.Remove(MarkedForDeath)
                    Me.Elements.Remove(MarkedForDeath)
                End If
            End If
        End Sub

        Public Sub Cut()
            If Not Me.SelectedIndex = Nothing Then
                Me.AddToClipboard(Me.Containers.Item(Me.SelectedIndex))
            End If
            Me.Delete()
        End Sub

        Public Sub Undo()
            If Me.blnUndoable Then
                Try
                    Dim Panel As Panel = CType(Me.Containers.Item(Me.UndoIndex), Panel)
                    Panel.Top = Me.UndoTop
                    Panel.Left = Me.UndoLeft
                    Panel.Width = Me.UndoWidth
                    Panel.Height = Me.UndoHeight
                Catch
                    Me.Paste(Me.UndoDelete)
                End Try
                Me.blnUndoable = False
            End If
        End Sub

        Public Sub SaveUndoStep(ByVal Index As String)
            Try
                If Not Index = Nothing And Not Me.Containers(Index) Is Nothing Then
                    Dim Panel As Panel = CType(Me.Containers.Item(Index), Panel)
                    If Not Panel Is Nothing Then
                        Me.UndoTop = Panel.Top
                        Me.UndoLeft = Panel.Left
                        Me.UndoWidth = Panel.Width
                        Me.UndoHeight = Panel.Height
                        Me.UndoIndex = Index
                    End If
                End If
            Catch
                MsgBox("Undo Error")
            End Try
        End Sub

#End Region

#Region "User Interactions"

        Private Sub Panel_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
            If Not sender.Tag = Me.SelectedIndex Then
                Me.SelectElement(sender)
            End If
        End Sub

        Private Sub Label_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
            If Not sender.Tag = Me.SelectedIndex Then
                Me.SelectElement(sender.Parent)
            End If
        End Sub

        Private Sub TextBox_Enter(ByVal sender As Object, ByVal e As System.EventArgs)
            Me.Content.Select()
        End Sub

        Private Sub Element_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
            If Me.DefaultAction = Action.Move Or Me.DefaultAction = Action.Resize Then
                If e.Button = MouseButtons.Left Then
                    Dim Div As Panel
                    If sender.GetType.ToString = "System.Windows.Forms.Panel" Then
                        Div = Containers(sender.Tag)
                    Else
                        Div = Containers(sender.Parent.Tag)
                    End If
                    If Me.blnDrag And Me.CanEdit Then
                        If DefaultAction = Action.Move Then
                            Dim MouseY As Integer = Div.Top + (e.Y - MoveY)
                            Dim MouseX As Integer = Div.Left + (e.X - MoveX)
                            If Me.Grid = 0 Then
                                Div.Top = MouseY
                                Div.Left = MouseX
                            Else
                                Div.Top = Math.Round(MouseY / Grid) * Grid
                                Div.Left = Math.Round(MouseX / Grid) * Grid
                            End If
                        ElseIf DefaultAction = Action.Resize Then
                            Dim MouseY As Integer = e.Y
                            Dim MouseX As Integer = e.X
                            If MouseX > 2 Then
                                If Me.Grid = 0 Then
                                    Div.Width = MouseX
                                Else
                                    Div.Width = Math.Round(MouseX / Grid) * Grid
                                End If
                            End If
                            If MouseY > 2 Then
                                If Me.Grid = 0 Then
                                    Div.Height = MouseY
                                Else
                                    Div.Height = Math.Round(MouseY / Grid) * Grid
                                End If
                            End If
                        End If
                    End If
                    RaiseEvent SelectedElement_MouseMove(sender, e)
                End If
            End If
        End Sub

        Private Sub Element_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
            If Me.DefaultAction = Action.Move Or Me.DefaultAction = Action.Resize Then
                If e.Button = MouseButtons.Left Then
                    Dim Div As Panel
                    If sender.GetType.ToString = "System.Windows.Forms.Panel" Then
                        Div = Containers(sender.Tag)
                    Else
                        Div = Containers(sender.Parent.Tag)
                    End If
                    Me.SelectElement(Div)
                    Me.blnDrag = True
                    MoveX = e.X
                    MoveY = e.Y
                    Me.SaveUndoStep(Me.SelectedIndex)
                    Me.blnUndoable = True
                    RaiseEvent DragStatusChanged(True)
                    RaiseEvent SelectedElement_MouseDown(Div, e)
                End If
            End If
        End Sub

        Private Sub Element_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
            If Me.DefaultAction = Action.Move Or Me.DefaultAction = Action.Resize Then
                If e.Button = MouseButtons.Left Then
                    Dim Div As Panel
                    If sender.GetType.ToString = "System.Windows.Forms.Panel" Then
                        Div = Containers(sender.Tag)
                    Else
                        Div = Containers(sender.Parent.Tag)
                    End If
                    Me.DragOff()
                    RaiseEvent SelectedElement_MouseUp(Div, e)
                End If
            End If
        End Sub

        Private Sub Element_DoubleClick(ByVal sender As Object, ByVal e As EventArgs)
            Me.SelectElement(sender.Parent)
            RaiseEvent SelectedElement_DoubleClick(sender, e)
        End Sub

        Private Sub Element_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
            If Me.DefaultAction = Action.Move Then
                If Not sender.Tag = Me.SelectedIndex Then
                    Me.SelectElement(sender.Parent)
                End If
                RaiseEvent SelectedElement_KeyPress(sender.Parent, e)
            End If
        End Sub

        Private Sub DragOff()
            Me.blnDrag = False
            RaiseEvent DragStatusChanged(False)
        End Sub

        Public Sub SelectElement(ByVal Element As Panel)
            Me.UnselectElement()
            SelectedIndex = Element.Tag
            CType(Me.Containers.Item(Me.SelectedIndex), Panel).BackColor = System.Drawing.Color.DarkRed
            RaiseEvent SelectedIndexChanged(Element)
        End Sub

        Public Sub UnselectElement()
            Try
                If Not Me.SelectedIndex = Nothing Then
                    CType(Me.Containers.Item(Me.SelectedIndex), Panel).BackColor = Me.DivOutline
                End If
            Catch ex As Exception
            End Try
            SelectedIndex = Nothing
            Me.DragOff()
            Me.Content.Focus()
            RaiseEvent SelectedIndexChanged(Nothing)
        End Sub



#End Region

    End Class


    Public Class HtmlDocument

        Dim _Template As String = ""
        Dim _Original As String = ""
        Dim _Report As New Xml.XmlDocument

        Public Property Template() As String
            Get
                Return _Template
            End Get
            Set(ByVal Value As String)
                Me._Template = Value
                Me._Report.LoadXml(Value)
            End Set
        End Property

        Public ReadOnly Property Report() As XmlDocument
            Get
                Return _Report
            End Get
        End Property

        Public Sub New(ByVal Template As String)
            Me._Template = Template
            Me._Original = Template
            Me._Report.LoadXml(Me._Template)
        End Sub

        Public Sub New()

        End Sub

        Public Sub Clear()
            Me._Template = Me._Original
            Me._Report.LoadXml(Me._Template)
        End Sub

        Public Overrides Function ToString() As String
            Return Me._Report.SelectSingleNode("/").OuterXml
        End Function

        Public Sub Replace(ByVal This As String, ByVal That As String)
            ' Revert if already escaped
            That = That.Replace("&amp;", "&")
            ' Now clean again
            That = That.Replace("&", "&amp;")
            ' Replace this with that
            Me._Template = Me._Template.Replace(This, That)
            ' Load into XML Doc
            Me._Report.LoadXml(Me._Template)
        End Sub

        Private Function FilterText(ByVal Text As String) As String
            Text = Text.Replace("&", "&amp;")
            Return Text
        End Function

        Public Sub PopulateTable(ByVal SourceName As String, ByVal Data As DataTable)
            ' Find nodes
            Dim Nodes As Xml.XmlNodeList
            Dim Node As Xml.XmlNode
            Dim Style As String = ""
            'Style = "font-size: 12px; font-family: Arial; "
            Nodes = Me._Report.SelectNodes("/html/body/page/div/table")
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
                            Th.Attributes.Append(Me._Report.CreateAttribute("style")).InnerText = Style
                        End If
                        i += 1
                    Next
                    ' If there is data in Data variable
                    If Data.Rows.Count > 0 Then
                        ' Add Tr For each row
                        For Each Row In Data.Rows
                            Dim Tr As XmlElement = Me._Report.CreateElement("tr")
                            Tr.Attributes.Append(Me._Report.CreateAttribute("valign")).InnerText = "top"
                            For i = 0 To Columns.Length - 1
                                Dim Td As XmlElement = Me._Report.CreateElement("td")
                                Try
                                    Td.InnerText = Format(IfNull(Row.Item(Columns(i)), ""), Formats(i))
                                Catch ex As Exception
                                    MsgBox("Column did not exist in table.  Template is probably wrong. Here is exact error: " & ex.ToString)
                                End Try
                                If Style.Length > 0 Then
                                    Td.Attributes.Append(Me._Report.CreateAttribute("style")).InnerText = Style
                                End If
                                Tr.AppendChild(Td)
                            Next
                            Node.AppendChild(Tr)
                        Next
                    Else
                        Dim Tr As XmlElement = Me._Report.CreateElement("tr")
                        Dim Td As XmlElement = Me._Report.CreateElement("td")
                        Dim ColSpan As XmlAttribute = Me._Report.CreateAttribute("colspan")
                        Td.Attributes.Append(ColSpan).InnerText = Node.SelectNodes("tr/th").Count.ToString
                        Td.InnerText = "None"
                        Tr.AppendChild(Td)
                        Node.AppendChild(Tr)
                    End If
                End If
            Next
        End Sub

        Private Function IfNull(ByVal Value As Object, Optional ByVal NewVal As String = "") As String
            If Value Is DBNull.Value Then
                Return NewVal
            Else
                Return Value
            End If
        End Function

        Public Sub MakeImagesExternal()
            Dim Nodes As Xml.XmlNodeList
            Dim Node As Xml.XmlNode
            Dim strKey As String
            Dim hashImage As New Hashtable
            Dim SrcPath As String
            Dim MD5 As New MD5CryptoServiceProvider
            ' Loop through and replace images
            Nodes = Me._Report.SelectNodes("//img/@src")
            For Each Node In Nodes
                If Node.InnerText.StartsWith("data:image/gif;base64,") Then
                    ' Save to a local image and replace src with path
                    Node.InnerText = Node.InnerText.Replace("data:image/gif;base64,", "")
                    strKey = BitConverter.ToString(MD5.ComputeHash(System.Text.UnicodeEncoding.ASCII.GetBytes(Node.InnerText)))
                    If Not hashImage.ContainsKey(strKey) Then
                        Dim MemStream As New IO.MemoryStream(Convert.FromBase64String(Node.InnerText))
                        Dim Image As New Drawing.Bitmap(MemStream)
                        SrcPath = Environment.GetEnvironmentVariable("TEMP") & "\gravity.img." & Now.Ticks & "-" & strKey.Substring(0, 8) & ".gif"
                        Image.Save(SrcPath, Drawing.Imaging.ImageFormat.Gif)
                        hashImage.Add(strKey, SrcPath)
                    Else
                        SrcPath = hashImage(strKey)
                    End If
                    Node.InnerText = "file:///" & SrcPath.Replace("\", "/")
                End If
            Next
        End Sub

    End Class

    Public Class TableRestraint
        Public Src As String = ""
        Public Start As Integer = 0
        Public Limit As Integer = 10
    End Class

End Namespace


