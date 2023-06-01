Namespace GravityDocument

    Public Class gElement

        Public ReadOnly Property IsRollover() As Boolean
            Get
                If Me.ClassName = Classes.Table Then
                    If Me.Table.RowsPerPage > 0 And Me.Table.Data IsNot Nothing Then
                        If Me.Table.Data.Rows.Count > Me.Table.RowsPerPage Then
                            Return True
                        End If
                    End If
                End If
                Return False
            End Get
        End Property

#Region "Structures/Enums"

        Public Enum Classes
            Label = 1
            Image = 2
            Table = 3
            Signature = 4
        End Enum

        Public Enum LineStyle
            None = 0
            Solid = 1
            Dashed = 2
            Dotted = 2
            DoubleLine = 3
        End Enum

        Public Enum HorizontalAlignment
            Left = 0
            Center = 1
            Right = 2
            Justify = 3
        End Enum

        Public Enum OverflowMethods
            Auto = 0
            Visible = 1
            Hidden = 2
        End Enum

#End Region

#Region "Private Variables"

        Dim _Class As Classes = Classes.Label
        Dim _Id As String = ""
        Dim _Key As String = ""
        Dim _Top As Integer = 0
        Dim _Left As Integer = 0
        Dim _Width As Integer = 100
        Dim _Height As Integer = 100
        Dim _Text As String = ""
        Dim _ReadOnly As Boolean = False
        Dim _Bold As Boolean = False
        Dim _Italic As Boolean = False
        Dim _Underlined As Boolean = False
        Dim _FontSize As Integer = 10
        Dim _FontFamily As String = "Arial"
        Dim _Bgcolor As System.Drawing.Color = Drawing.Color.Transparent
        Dim _Fgcolor As System.Drawing.Color = Drawing.Color.Black
        Dim _Padding As Integer = 0
        Dim _BorderColor As System.Drawing.Color = Drawing.Color.Black
        Dim _BorderWidth As Integer = 0
        Dim _BorderStyle As LineStyle = LineStyle.None
        Dim _TextAlign As HorizontalAlignment = HorizontalAlignment.Left
        Dim _Overflow As OverflowMethods = OverflowMethods.Auto
        Dim _ZIndex As Integer = 0
        Dim _Page As gPage

#End Region

#Region "Public Variables"

        Public WithEvents Image As New gImage(Me)
        Public WithEvents Table As New gTable(Me)
        Public OriginalSource As String

#End Region

#Region "Events"

        Public Event LocationChanged(ByVal sender As gElement, ByVal x As Integer, ByVal y As Integer)
        Public Event SizeChanged(ByVal sender As gElement, ByVal x As Integer, ByVal y As Integer)
        Public Event TextChanged(ByVal sender As gElement, ByVal Value As String)
        Public Event ReadOnlyChanged(ByVal sender As gElement, ByVal RO As Boolean)
        Public Event FontStyleChanged(ByVal sender As gElement, ByVal Family As String, ByVal Size As Integer, ByVal Bold As Boolean, ByVal Italic As Boolean, ByVal Underlined As Boolean)
        Public Event ForeColorChanged(ByVal sender As gElement, ByVal fg As System.Drawing.Color)
        Public Event BackColorChanged(ByVal sender As gElement, ByVal bg As System.Drawing.Color)
        Public Event PaddingChanged(ByVal sender As gElement, ByVal Value As Integer)
        Public Event BorderChanged(ByVal sender As gElement, ByVal Style As LineStyle, ByVal Width As Integer, ByVal Color As System.Drawing.Color)
        Public Event TextAlignChanged(ByVal sender As gElement, ByVal al As HorizontalAlignment)
        Public Event OverflowChanged(ByVal sender As gElement, ByVal mode As OverflowAction)
        Public Event ZIndexChanged(ByVal sender As gElement, ByVal z As Integer)
        Public Event ImageChanged(ByVal sender As gElement, ByVal image As gImage)
        Public Event IdChanged(ByVal sender As gElement, ByVal Id As String)

#End Region

#Region "Properties"

        Public ReadOnly Property Page() As gPage
            Get
                Return Me._Page
            End Get
        End Property


        Public ReadOnly Property ClassName() As Classes
            Get
                Return Me._Class
            End Get
        End Property

        Public Property Id() As String
            Get
                Return Me._Id
            End Get
            Set(ByVal value As String)
                Me._Id = value
                RaiseEvent IdChanged(Me, value)
            End Set
        End Property

        Public ReadOnly Property Key() As String
            Get
                Return Me._Key
            End Get
        End Property

        Public Property Top() As Integer
            Get
                Return Me._Top
            End Get
            Set(ByVal Value As Integer)
                Me._Top = Value
                RaiseEvent LocationChanged(Me, Me._Left, Me._Top)
            End Set
        End Property

        Public Property Left() As Integer
            Get
                Return Me._Left
            End Get
            Set(ByVal Value As Integer)
                Me._Left = Value
                RaiseEvent LocationChanged(Me, Me._Left, Me._Top)
            End Set
        End Property

        Public Property Width() As Integer
            Get
                Return Me._Width
            End Get
            Set(ByVal Value As Integer)
                Me._Width = Value
                RaiseEvent SizeChanged(Me, Me._Width, Me._Height)
            End Set
        End Property

        Public Property Height() As Integer
            Get
                Return Me._Height
            End Get
            Set(ByVal Value As Integer)
                Me._Height = Value
                RaiseEvent SizeChanged(Me, Me._Width, Me._Height)
            End Set
        End Property

        Public Property Text() As String
            Get
                Return Me._Text
            End Get
            Set(ByVal value As String)
                value = Me.Page.Document.TextClean(value)
                Me._Text = value
                RaiseEvent TextChanged(Me, value)
            End Set
        End Property

        Public Property Bgcolor() As System.Drawing.Color
            Get
                Return Me._Bgcolor
            End Get
            Set(ByVal value As System.Drawing.Color)
                Me._Bgcolor = value
                RaiseEvent BackColorChanged(Me, value)
            End Set
        End Property

        Public Property Fgcolor() As System.Drawing.Color
            Get
                Return Me._Fgcolor
            End Get
            Set(ByVal value As System.Drawing.Color)
                Me._Fgcolor = value
                RaiseEvent ForeColorChanged(Me, value)
            End Set
        End Property

        Public Property FontSize() As Integer
            Get
                Return Me._FontSize
            End Get
            Set(ByVal Value As Integer)
                If Value < 1 Then
                    Me._FontSize = 1
                ElseIf Value > 100 Then
                    Me._FontSize = 100
                Else
                    Me._FontSize = Value
                End If
                RaiseEvent FontStyleChanged(Me, Me.FontFamily, Me.FontSize, Me.FontBold, Me.FontItalic, Me.FontUnderline)
            End Set
        End Property

        Public Property FontBold() As Boolean
            Get
                Return Me._Bold
            End Get
            Set(ByVal Value As Boolean)
                Me._Bold = Value
                RaiseEvent FontStyleChanged(Me, Me.FontFamily, Me.FontSize, Me.FontBold, Me.FontItalic, Me.FontUnderline)
            End Set
        End Property

        Public Property FontItalic() As Boolean
            Get
                Return Me._Italic
            End Get
            Set(ByVal Value As Boolean)
                Me._Italic = Value
                RaiseEvent FontStyleChanged(Me, Me.FontFamily, Me.FontSize, Me.FontBold, Me.FontItalic, Me.FontUnderline)
            End Set
        End Property

        Public Property FontUnderline() As Boolean
            Get
                Return Me._Underlined
            End Get
            Set(ByVal Value As Boolean)
                Me._Underlined = Value
                RaiseEvent FontStyleChanged(Me, Me.FontFamily, Me.FontSize, Me.FontBold, Me.FontItalic, Me.FontUnderline)
            End Set
        End Property

        Public Property FontFamily() As String
            Get
                Return Me._FontFamily
            End Get
            Set(ByVal value As String)
                Me._FontFamily = value
                RaiseEvent FontStyleChanged(Me, Me.FontFamily, Me.FontSize, Me.FontBold, Me.FontItalic, Me.FontUnderline)
            End Set
        End Property

        Public Property TextAlign() As HorizontalAlignment
            Get
                Return Me._TextAlign
            End Get
            Set(ByVal value As HorizontalAlignment)
                Me._TextAlign = value
                RaiseEvent TextAlignChanged(Me, value)
            End Set
        End Property

        Public Property BorderWidth() As Integer
            Get
                Return Me._BorderWidth
            End Get
            Set(ByVal value As Integer)
                Me._BorderWidth = value
                RaiseEvent BorderChanged(Me, Me._BorderStyle, Me._BorderWidth, Me._BorderColor)
            End Set
        End Property

        Public Property BorderColor() As Drawing.Color
            Get
                Return Me._BorderColor
            End Get
            Set(ByVal value As Drawing.Color)
                Me._BorderColor = value
                RaiseEvent BorderChanged(Me, Me._BorderStyle, Me._BorderWidth, Me._BorderColor)
            End Set
        End Property

        Public Property BorderStyle() As LineStyle
            Get
                Return Me._BorderStyle
            End Get
            Set(ByVal value As LineStyle)
                Me._BorderStyle = value
                RaiseEvent BorderChanged(Me, Me._BorderStyle, Me._BorderWidth, Me._BorderColor)
            End Set
        End Property

        Public Property Overflow() As OverflowAction
            Get
                Return Me._Overflow
            End Get
            Set(ByVal value As OverflowAction)
                Me._Overflow = value
                RaiseEvent OverflowChanged(Me, value)
            End Set
        End Property

        Public Property ZIndex() As Integer
            Get
                Return Me._ZIndex
            End Get
            Set(ByVal value As Integer)
                Me._ZIndex = value
                RaiseEvent ZIndexChanged(Me, value)
            End Set
        End Property

#End Region

#Region "Methods"

        Public Sub New(ByRef Page As gPage, ByVal Key As String, Optional ByVal cn As Classes = Classes.Label)
            Me._Page = Page
            Me._Class = cn
            Me._Key = Key
            Me._Id = Key
        End Sub

        Friend Sub ChangeKey(ByVal Value As String)
            Me._Key = Value
        End Sub

        Public Sub Remove()
            Me.Page.RemoveElement(Me.Key)
        End Sub

        Private Function StripNonAlphaNumeric(ByVal Input As String) As String
            Dim str As String = "[^a-zA-Z0-9]"
            Dim regex As New System.Text.RegularExpressions.Regex(str)
            Dim out As String = regex.Replace(Input, "")
            Return out
        End Function

        Public Function ToXml(Optional ByVal InternalImages As Boolean = True, Optional ByVal MoveDown As Integer = 0, Optional ByVal Export As Boolean = False) As String
            Dim Xml As String = ""
            Xml &= "<div id=""" & Me.Id & """ class=""" & Me.ClassName.ToString & """"
            Xml &= " style=""" & Me.GetStyleString(MoveDown) & """>"
            Select Case Me.ClassName
                Case Classes.Label
                    Xml &= Me.Page.Document.XmlClean(Me.Text)
                Case Classes.Image, Classes.Signature
                    Xml &= "<img src="""
                    If InternalImages Then
                        Xml &= Me.Image.ToDataUrl
                    Else
                        Dim Key As String = Me.Image.ToString.Substring(0, 32)
                        If Not Me.Page.Document.LocalImages.ContainsKey(Key) Then
                            Dim Bmp As Drawing.Bitmap = Me.Image.ToBmp
                            Dim Path As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\gravity." & Now.Ticks & Me.StripNonAlphaNumeric(Key) & ".gif"
                            If Bmp IsNot Nothing Then
                                Bmp.Save(Path, Drawing.Imaging.ImageFormat.Gif)
                                Me.Page.Document.LocalImages.Add(Key, Path)
                                Xml &= Path
                            End If
                        Else
                            Xml &= Me.Page.Document.LocalImages(Key)
                        End If
                    End If
                    Xml &= """ alt=""" & Me.Page.Document.XmlClean(Me.Text) & """ />"
                Case Classes.Table
                    Xml &= "<table"
                    Xml &= " src=""" & Me.Table.Source & """"
                    Xml &= " rowlimit=""" & Me.Table.RowsPerPage & """"
                    Xml &= " cellspacing=""" & Me.Table.CellSpacing & """"
                    Xml &= " cellpadding=""" & Me.Table.CellPadding & """"
                    If Export Then
                        Xml &= " style=""" & Me.GetStyleString(0, True) & """"
                    End If
                    Xml &= ">"
                    Xml &= "<tr>"
                    For Each col As gTable.Column In Me.Table.Columns
                        Xml &= "<th map=""" & col.Key & """ format=""" & col.Format & """ width=""" & col.Width & """ align=""" & col.Align & """>"
                        Xml &= Me.Page.Document.XmlClean(col.HeaderText) & "</th>"
                    Next
                    Xml &= "</tr>"
                    Dim Count As Integer = 1
                    For Each r As DataRow In Me.Table.Data.Rows
                        ' Only print up to max (intentionally one less if it's full to make room for continued row)
                        If Me.Table.RowsPerPage = 0 Or Count < Me.Table.RowsPerPage Or Me.Table.Data.Rows.Count = Me.Table.RowsPerPage Then
                            Dim tr As String = ""
                            tr &= "<tr>"
                            For Each col As gTable.Column In Me.Table.Columns
                                Dim Value As String = IIf(r.Item(col.Key) Is DBNull.Value, "", r.Item(col.Key))
                                tr &= "<td align=""" & col.Align & """>"
                                Try
                                    If col.Format.Trim.Length = 0 Then
                                        tr &= Me.Page.Document.XmlClean(Value)
                                    ElseIf col.Format = "percent" Then
                                        tr &= Me.Page.Document.XmlClean(Math.Round(CDbl(Value), 3) & "%")
                                    Else
                                        tr &= Me.Page.Document.XmlClean(Format(Value, col.Format))
                                    End If
                                Catch ex As Exception
                                    tr &= Me.Page.Document.XmlClean(Value)
                                End Try
                                tr &= "</td>"
                            Next
                            tr &= "</tr>"
                            Xml &= tr
                        Else
                            Exit For
                        End If
                        Count += 1
                    Next
                    If Me.IsRollover Then
                        ' Add one more row to this one to say it's continued
                        Xml &= "<tr><td colspan=""" & Me.Table.Columns.Count & """>Continued on next page...</td></tr>"
                    End If
                    Xml &= "</table>"
            End Select
            Xml &= "</div>"
            Return Xml
        End Function

        Public Function GetStyleString(Optional ByVal MoveDown As Integer = 0, Optional ByVal IsSubElement As Boolean = False, Optional ByVal AllowBorder As Boolean = True) As String
            Dim Style As String = ""
            If Not IsSubElement Then
                Style = "position: absolute; "
                Style &= " left: " & Me.Left & "px; "
                Style &= " top: " & Me.Top + MoveDown & "px; "
                Style &= " width: " & Me.Width & "px; "
                Style &= " height: " & Me.Height & "px; "
                Style &= " z-index: " & Me.ZIndex & "; "
                If Me._Bgcolor <> Drawing.Color.Transparent Then
                    Style &= "background-color: rgb(" & Me._Bgcolor.R & ", " & Me._Bgcolor.G & ", " & Me._Bgcolor.B & "); "
                End If
            End If
            If AllowBorder Then
                If Me._BorderStyle = LineStyle.None Then
                    Style &= "border-style: none; "
                Else
                    Style &= "border-width: " & Me._BorderWidth & "; "
                    Style &= "border-color: rgb(" & Me._BorderColor.R & ", " & Me._BorderColor.G & ", " & Me._BorderColor.B & "); "
                    If Me._BorderStyle = LineStyle.Dashed Then
                        Style &= "border-style: dashed; "
                    ElseIf Me._BorderStyle = LineStyle.Dotted Then
                        Style &= "border-style: dotted; "
                    ElseIf Me._BorderStyle = LineStyle.DoubleLine Then
                        Style &= "border-style: double; "
                    Else
                        Style &= "border-style: solid; "
                    End If
                End If
            End If
            Style &= " color: rgb(" & Me._Fgcolor.R & ", " & Me._Fgcolor.G & ", " & Me._Fgcolor.B & "); "
            Style &= " font-weight: " & IIf(Me._Bold, "bold", "normal") & "; "
            Style &= " font-family: '" & Me._FontFamily & "'; "
            Style &= " font-size: " & Me._FontSize & "pt; "
            Style &= " text-decoration: " & IIf(Me._Underlined, "underline", "none") & "; "
            Style &= " font-style: " & IIf(Me._Italic, "italic", "normal") & "; "
            Style &= " text-align: "
            If Me._TextAlign = HorizontalAlignment.Center Then
                Style &= "center; "
            ElseIf Me._TextAlign = HorizontalAlignment.Left Then
                Style &= "left; "
            ElseIf Me._TextAlign = HorizontalAlignment.Justify Then
                Style &= "justify; "
            Else
                Style &= "right; "
            End If
            ' Wrapping
            'If Me.ClassName = Classes.Label Then
            '    Style &= "white-space: pre-wrap; "
            'End If
            Return Style
        End Function

#End Region

#Region "Child Events"

        Private Sub Image_ImageChanged(ByVal sender As gImage) Handles Image.ImageChanged
            RaiseEvent ImageChanged(Me, sender)
        End Sub

#End Region

    End Class

End Namespace
