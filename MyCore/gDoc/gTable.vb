Namespace GravityDocument

    Public Class gTable

        Dim _Data As New DataTable
        Dim _Source As String
        Dim _Padding As Integer = 0
        Dim _Spacing As Integer = 0
        Dim _RowsPerPage As Integer = 0
        Dim _Element As gElement
        Dim _Columns As New Collection

        Public Event ColumnsCleared(ByVal sender As gTable)
        Public Event ColumnAdded(ByVal sender As gTable, ByVal col As Column)
        Public Event DataSourceChanged(ByVal sender As gTable)
        Public Event SourceChanged(ByVal sender As gTable, ByVal Source As String)
        Public Event CellPaddingChanged(ByVal sender As gTable, ByVal Padding As Integer)
        Public Event CellSpacingChanged(ByVal sender As gTable, ByVal Spacing As Integer)
        Public Event RowsPerPageChanged(ByVal sender As gTable, ByVal Rows As Integer)
        Public Event ColumnWidthChanged(ByVal col As Column, ByVal NewWidth As Integer)

        Public Property Data() As DataTable
            Get
                Return Me._Data
            End Get
            Set(ByVal value As DataTable)
                If value IsNot Me._Data Then
                    Me._Data = value
                    RaiseEvent DataSourceChanged(Me)
                End If
            End Set
        End Property

        Public Property Source() As String
            Get
                Return Me._Source
            End Get
            Set(ByVal value As String)
                If value <> Me._Source Then
                    Me._Source = value
                    RaiseEvent SourceChanged(Me, value)
                End If
            End Set
        End Property

        Public Property CellPadding() As Integer
            Get
                Return Me._Padding
            End Get
            Set(ByVal value As Integer)
                If value <> Me._Padding Then
                    Me._Padding = value
                    RaiseEvent CellPaddingChanged(Me, value)
                End If
            End Set
        End Property

        Public Property CellSpacing() As Integer
            Get
                Return Me._Spacing
            End Get
            Set(ByVal value As Integer)
                If value <> Me._Spacing Then
                    Me._Spacing = value
                    RaiseEvent CellSpacingChanged(Me, value)
                End If
            End Set
        End Property

        Public Property RowsPerPage() As Integer
            Get
                Return Me._RowsPerPage
            End Get
            Set(ByVal value As Integer)
                If value <> Me._RowsPerPage Then
                    Me._RowsPerPage = value
                    RaiseEvent RowsPerPageChanged(Me, value)
                End If
            End Set
        End Property

        Public ReadOnly Property Element() As gElement
            Get
                Return Me._Element
            End Get
        End Property

        Public Sub New(ByVal ParentElement As gElement)
            Me._Element = ParentElement
        End Sub

        Public Function Columns(ByVal Index As Integer) As Column
            Return Me._Columns(Index)
        End Function

        Public Function Columns(ByVal Key As String) As Column
            Return Me._Columns(Key)
        End Function

        Public Function Columns() As Collection
            Return Me._Columns
        End Function

        Public Sub ClearColumns()
            Me._Columns.Clear()
            RaiseEvent ColumnsCleared(Me)
        End Sub

        Public Sub AddColumn(ByVal Col As Column, Optional ByVal Key As String = Nothing)
            Me._Columns.Add(Col, Key)
            AddHandler Col.WidthChanged, AddressOf ColWidthChanged
            RaiseEvent ColumnAdded(Me, Col)
        End Sub

        Private Sub ColWidthChanged(ByVal c As Column, ByVal x As Integer)
            RaiseEvent ColumnWidthChanged(c, x)
        End Sub

        Public Class Column

            Dim _Key As String
            Dim _Width As Integer
            Dim _HeaderText As String
            Dim _Format As String
            Dim _Align As String
            Dim _Parent As gTable

            Public Event WidthChanged(ByVal sender As Column, ByVal NewWidth As Integer)
            Public Event HeaderTextChanged(ByVal sender As Column, ByVal NewText As String)
            Public Event KeyChanged(ByVal sender As Column, ByVal NewKey As String)
            Public Event FormatChanged(ByVal sender As Column, ByVal NewFormat As String)
            Public Event AlignmentChanged(ByVal sender As Column, ByVal NewAlignment As String)

            Public Property Key() As String
                Get
                    Return Me._Key
                End Get
                Set(ByVal value As String)
                    If Me._Key <> value Then
                        Me._Key = value
                        RaiseEvent KeyChanged(Me, value)
                    End If
                End Set
            End Property

            Public Property HeaderText() As String
                Get
                    Return Me._HeaderText
                End Get
                Set(ByVal value As String)
                    If Me._HeaderText <> value Then
                        Me._HeaderText = value
                        RaiseEvent HeaderTextChanged(Me, value)
                    End If
                End Set
            End Property

            Public Property Format() As String
                Get
                    Return Me._Format
                End Get
                Set(ByVal value As String)
                    If Me._Format <> value Then
                        Me._Format = value
                        RaiseEvent FormatChanged(Me, value)
                    End If
                End Set
            End Property

            Public Property Align() As String
                Get
                    Return Me._Align.ToLower
                End Get
                Set(ByVal value As String)
                    If Me._Align <> value Then
                        Me._Align = value
                        RaiseEvent AlignmentChanged(Me, value)
                    End If
                End Set
            End Property

            Public Property Width() As Integer
                Get
                    Return Me._Width
                End Get
                Set(ByVal value As Integer)
                    If Me._Width <> value Then
                        Me._Width = value
                        RaiseEvent WidthChanged(Me, value)
                    End If
                End Set
            End Property

            Public ReadOnly Property Table() As gTable
                Get
                    Return Me._Parent
                End Get
            End Property

            Public Sub New(ByVal ParentTable As gTable)
                Me._Parent = ParentTable
            End Sub


        End Class

    End Class

End Namespace
