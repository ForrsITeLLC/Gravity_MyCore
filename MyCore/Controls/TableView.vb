Public Class TableView
    Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'UserControl overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents ListView As System.Windows.Forms.ListView
    Friend WithEvents BottomPanel As System.Windows.Forms.Panel
    Friend WithEvents BottomText As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.BottomPanel = New System.Windows.Forms.Panel
        Me.BottomText = New System.Windows.Forms.TextBox
        Me.ListView = New System.Windows.Forms.ListView
        Me.BottomPanel.SuspendLayout()
        Me.SuspendLayout()
        '
        'BottomPanel
        '
        Me.BottomPanel.Controls.Add(Me.BottomText)
        Me.BottomPanel.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.BottomPanel.Location = New System.Drawing.Point(0, 272)
        Me.BottomPanel.Name = "BottomPanel"
        Me.BottomPanel.Size = New System.Drawing.Size(416, 48)
        Me.BottomPanel.TabIndex = 0
        '
        'BottomText
        '
        Me.BottomText.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.BottomText.Dock = System.Windows.Forms.DockStyle.Fill
        Me.BottomText.Location = New System.Drawing.Point(0, 0)
        Me.BottomText.Multiline = True
        Me.BottomText.Name = "BottomText"
        Me.BottomText.Size = New System.Drawing.Size(416, 48)
        Me.BottomText.TabIndex = 0
        Me.BottomText.Text = ""
        Me.BottomText.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ListView
        '
        Me.ListView.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ListView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListView.Location = New System.Drawing.Point(0, 0)
        Me.ListView.Name = "ListView"
        Me.ListView.Size = New System.Drawing.Size(416, 272)
        Me.ListView.TabIndex = 1
        '
        'TableView
        '
        Me.Controls.Add(Me.ListView)
        Me.Controls.Add(Me.BottomPanel)
        Me.Name = "TableView"
        Me.Size = New System.Drawing.Size(416, 320)
        Me.BottomPanel.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public WithEvents Columns As New ColumnsCollection
    Public WithEvents Rows As New RowsCollection
    Public Shadows Event Enter(ByVal sender As Object, ByVal e As System.EventArgs)

    Dim _ds As String = Nothing
    Dim _rowlimit As Integer = 0
    Dim _cellpadding As Integer = 1
    Dim _cellspacing As Integer = 1
    Dim _cellfontsize As Integer = 8
    Dim _cellfontname As String = "Arial"

    Public Property ShowCaption() As Boolean
        Get
            Return Me.BottomPanel.Visible
        End Get
        Set(ByVal Value As Boolean)
            Me.BottomPanel.Visible = Value
        End Set
    End Property

    Public Property CaptionSide() As CaptionEnum
        Get
            If Me.BottomPanel.Dock = DockStyle.Bottom Then
                Return CaptionEnum.Bottom
            Else
                Return CaptionEnum.Top
            End If
        End Get
        Set(ByVal Value As CaptionEnum)
            If Value = CaptionEnum.Bottom Then
                Me.BottomPanel.Dock = DockStyle.Bottom
            Else
                Me.BottomPanel.Dock = DockStyle.Top
            End If
        End Set
    End Property

    Public Property CaptionPadding() As Integer
        Get
            Return BottomPanel.DockPadding.All
        End Get
        Set(ByVal Value As Integer)
            BottomPanel.DockPadding.All = Value
        End Set
    End Property

    Public Property CaptionText() As String
        Get
            Return BottomText.Text
        End Get
        Set(ByVal Value As String)
            BottomText.Text = Value
        End Set
    End Property

    Public Property CaptionFont() As Font
        Get
            Return Me.BottomText.Font
        End Get
        Set(ByVal Value As Font)
            Me.BottomText.Font = Value
        End Set
    End Property

    Public Property CaptionTextAlign() As HorizontalAlignment
        Get
            Return BottomText.TextAlign
        End Get
        Set(ByVal Value As HorizontalAlignment)
            BottomText.TextAlign = Value
        End Set
    End Property

    Public Property CaptionBackColor() As Color
        Get
            Return BottomPanel.BackColor
        End Get
        Set(ByVal Value As Color)
            BottomText.BackColor = Value
            BottomPanel.BackColor = Value
        End Set
    End Property

    Public Property CaptionForeColor() As Color
        Get
            Return BottomText.ForeColor
        End Get
        Set(ByVal Value As Color)
            BottomText.ForeColor = Value
        End Set
    End Property

    Public Property CaptionLocked() As Boolean
        Get
            Return BottomText.ReadOnly
        End Get
        Set(ByVal Value As Boolean)
            BottomText.ReadOnly = Value
        End Set
    End Property

    Public Property CaptionVisble() As Boolean
        Get
            Return Me.BottomPanel.Visible
        End Get
        Set(ByVal Value As Boolean)
            Me.BottomPanel.Visible = Value
        End Set
    End Property

    Public Property CaptionSize() As Integer
        Get
            Return Me.BottomPanel.Height
        End Get
        Set(ByVal Value As Integer)
            Me.BottomPanel.Height = Value
        End Set
    End Property

    Public Property BorderStyle() As BorderStyle
        Get
            Return ListView.BorderStyle
        End Get
        Set(ByVal Value As BorderStyle)
            ListView.BorderStyle = Value
            Me.BottomText.BorderStyle = Value
        End Set
    End Property

    Public Property FullRowSelect() As Boolean
        Get
            Return ListView.FullRowSelect
        End Get
        Set(ByVal Value As Boolean)
            ListView.FullRowSelect = Value
        End Set
    End Property

    Public Property View() As View
        Get
            Return ListView.View
        End Get
        Set(ByVal Value As View)
            ListView.View = Value
        End Set
    End Property

    Public Property HeaderStyle() As ColumnHeaderStyle
        Get
            Return ListView.HeaderStyle
        End Get
        Set(ByVal Value As ColumnHeaderStyle)
            ListView.HeaderStyle = Value
        End Set
    End Property

    Public ReadOnly Property Table() As ListView
        Get
            Return ListView
        End Get
    End Property

    Public Property DataSource() As String
        Get
            Return _ds
        End Get
        Set(ByVal Value As String)
            _ds = Value
        End Set
    End Property

    Public Property RowLimit() As Integer
        Get
            Return Me._rowlimit
        End Get
        Set(ByVal Value As Integer)
            Me._rowlimit = Value
        End Set
    End Property

    Public Property CellSpacing() As Integer
        Get
            Return Me._cellspacing
        End Get
        Set(ByVal Value As Integer)
            Me._cellspacing = Value
        End Set
    End Property

    Public Property CellPadding() As Integer
        Get
            Return Me._cellpadding
        End Get
        Set(ByVal Value As Integer)
            Me._cellpadding = Value
        End Set
    End Property

    Public Property CellFont() As String
        Get
            Return Me._cellfontname
        End Get
        Set(ByVal Value As String)
            Me._cellfontname = Value
        End Set
    End Property

    Public Property CellFontSize() As Integer
        Get
            Return Me._cellfontsize
        End Get
        Set(ByVal Value As Integer)
            Me._cellfontsize = Value
        End Set
    End Property

    Public Enum CaptionEnum
        Top
        Bottom
    End Enum

    Private Sub NewCol(ByRef Col As ColumnHeader) Handles Columns.NewCol
        Me.ListView.Columns.Add(Col)
    End Sub

    Private Sub NewRow(ByRef Row As Row) Handles Rows.NewRow
        Dim Fields(Row.Fields.Count) As String
        For i As Integer = 1 To Row.Fields.Count
            Fields(i - 1) = Row.Fields(i).ToString
        Next
        Me.Table.Items.Add(New ListViewItem(Fields))
    End Sub

    Private Sub ClearRows() Handles Rows.ClearRows
        Me.Table.Items.Clear()
    End Sub

    Private Sub RemoveCol(ByVal Col As ColumnHeader)
        Me.Table.Columns.Remove(Col)
    End Sub


#Region "Columns"

    Public Class ColumnsCollection

        Public Items As New Collection
        Public Event NewCol(ByRef Col As ColumnHeader)
        Public Event RemoveCol(ByRef Col As ColumnHeader)

        Public Sub New()

        End Sub

        Public Sub Add(ByVal Text As String, ByVal Width As Integer, Optional ByVal MappingName As String = Nothing, Optional ByVal Format As String = "")
            ' Create new ColumnHeader
            Dim Header As New ColumnHeader
            Header.Text = Text
            Header.Width = Width
            ' Raise event to add it to listview
            RaiseEvent NewCol(Header)
            ' New column
            Dim Col As New Column(Header, MappingName, Format)
            ' Add it to collection
            Items.Add(Col)
        End Sub

        Public Sub Remove(ByVal i As Integer)
            RaiseEvent RemoveCol(CType(Me.Items(i), Column).Header)
            Me.Items.Remove(i)
        End Sub

    End Class

    Public Class Column

        Public MappingName As String = ""
        Public Format As String = ""
        Public Header As ColumnHeader

        Public Sub New(ByRef Col As ColumnHeader, ByVal mn As String, Optional ByVal f As String = "")
            Me.Header = Col
            Me.MappingName = mn
            Me.Format = f
        End Sub

    End Class

#End Region

#Region "Rows"

    Public Class RowsCollection

        Public Items As New Collection
        Public Event NewRow(ByRef Row As Row)
        Public Event ClearRows()

        Public Sub New()

        End Sub

        Public Sub Add(ByVal Fields As Collection)
            Dim NewRow As New Row(Fields)
            RaiseEvent NewRow(NewRow)
            Items.Add(NewRow)
        End Sub

        Public Sub Clear()
            Items = New Collection
        End Sub

    End Class

    Public Class Row

        Public Fields As New Collection

        Public Sub New(ByVal Fields As Collection)
            Me.Fields = Fields
        End Sub

    End Class

#End Region

End Class
