Public Class LinkList
    Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.DataSource.Columns.Add("name")
        Me.DataSource.Columns.Add("value")

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
    Friend WithEvents Panel As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Panel = New System.Windows.Forms.Panel
        Me.SuspendLayout()
        '
        'Panel
        '
        Me.Panel.AutoScroll = True
        Me.Panel.BackColor = System.Drawing.Color.White
        Me.Panel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel.Location = New System.Drawing.Point(0, 0)
        Me.Panel.Name = "Panel"
        Me.Panel.Size = New System.Drawing.Size(268, 228)
        Me.Panel.TabIndex = 0
        '
        'LinkList
        '
        Me.Controls.Add(Me.Panel)
        Me.Name = "LinkList"
        Me.Size = New System.Drawing.Size(268, 228)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim WithEvents _Table As DataTable = New DataTable
    Dim _DisplayMember As String = "name"
    Dim _ValueMember As String = "value"
    Dim _PaddingLeft As Integer = 0
    Dim _PaddingTop As Integer = 0
    Dim _PaddingRight As Integer = 0
    Dim _PaddingBottom As Integer = 0
    Dim _PaddingAll As Integer = 0
    Dim _LineHeight As Integer = 20
    Dim _BackColor As Color
    Dim _HoverColor As Color = System.Drawing.Color.Red
    Dim _LinkColor As Color = System.Drawing.Color.Blue

    Public Event LinkClicked(ByVal Value As String)

    Public Sub SetDataSource(ByVal Table As DataTable, ByVal ValueMember As String, ByVal DisplayMember As String)
        Me._Table = Table
        Me._DisplayMember = DisplayMember
        Me._ValueMember = ValueMember
        Me.RedrawLinks()
    End Sub

    Public Property DataSource() As DataTable
        Get
            Return _Table
        End Get
        Set(ByVal Value As DataTable)
            Me._Table = Value
            Me.RedrawLinks()
        End Set
    End Property

    Public Property DisplayMember() As String
        Get
            Return Me._DisplayMember
        End Get
        Set(ByVal Value As String)
            Me._DisplayMember = Value
            Me.RedrawLinks()
        End Set
    End Property

    Public Property ValueMember() As String
        Get
            Return Me._ValueMember
        End Get
        Set(ByVal Value As String)
            Me._ValueMember = Value
            Me.RedrawLinks()
        End Set
    End Property

    Public Property PaddingTop() As Integer
        Get
            Return Me._PaddingTop
        End Get
        Set(ByVal Value As Integer)
            Me._PaddingTop = Value
            Me.RedrawLinks()
        End Set
    End Property

    Public Property PaddingLeft() As Integer
        Get
            Return Me._PaddingLeft
        End Get
        Set(ByVal Value As Integer)
            Me._PaddingLeft = Value
            Me.RedrawLinks()
        End Set
    End Property

    Public Property PaddingRight() As Integer
        Get
            Return Me._PaddingRight
        End Get
        Set(ByVal Value As Integer)
            Me._PaddingRight = Value
            Me.Panel.AutoScrollMargin = New System.Drawing.Size(Value, Me._PaddingBottom)
        End Set
    End Property

    Public Property PaddingBottom() As Integer
        Get
            Return Me._PaddingBottom
        End Get
        Set(ByVal Value As Integer)
            Me._PaddingBottom = Value
            Me.Panel.AutoScrollMargin = New System.Drawing.Size(Me._PaddingRight, Value)
        End Set
    End Property

    Public Property HoverColor() As Drawing.Color
        Get
            Return Me._HoverColor
        End Get
        Set(ByVal Value As Drawing.Color)
            Me._HoverColor = Value
            Me.RedrawLinks()
        End Set
    End Property

    Public Property LinkColor() As Drawing.Color
        Get
            Return Me._LinkColor
        End Get
        Set(ByVal Value As Drawing.Color)
            Me._LinkColor = Value
            Me.RedrawLinks()
        End Set
    End Property

    Public Property PaddingAll() As Integer
        Get
            Return Me._PaddingAll
        End Get
        Set(ByVal Value As Integer)
            Me._PaddingAll = Value
            Me._PaddingBottom = Value
            Me._PaddingLeft = Value
            Me._PaddingRight = Value
            Me._PaddingTop = Value
            Me.RedrawLinks()
        End Set
    End Property

    Public Shadows Property BackColor() As Color
        Get
            Return Me._BackColor
        End Get
        Set(ByVal Value As Color)
            Me._BackColor = Value
            Me.Panel.BackColor = Value
        End Set
    End Property

    Public Function AddLink(ByVal Value As String, ByVal Display As String) As DataRow
        Dim Row As DataRow = Me.Datasource.NewRow
        Row.Item(Me.DisplayMember) = Display
        Row.Item(Me.ValueMember) = Value
        Me.Datasource.Rows.Add(Row)
        Me.RedrawLinks()
        Return Row
    End Function

    Public Function GetValues() As String()
        Dim Values(Me.Datasource.Rows.Count - 1) As String
        For i As Integer = 0 To Me.Datasource.Rows.Count - 1
            Values(i) = Me.Datasource.Rows(i).Item(Me.ValueMember)
        Next
        Return Values
    End Function

    Private Sub RedrawLinks()
        Me.Panel.Controls.Clear()
        If Not Me._Table Is Nothing Then
            Dim Row As DataRow
            Dim Top As Integer = Me._PaddingTop
            For Each Row In Me._Table.Rows
                Dim Link As New LinkLabel
                Dim Width As Integer
                Me.Panel.Controls.Add(Link)
                If Not Me._DisplayMember = Nothing Then
                    Link.Text = Row.Item(Me._DisplayMember)
                ElseIf Not Me._ValueMember = Nothing Then
                    Link.Text = Row.Item(Me._ValueMember)
                End If
                If Not Me._ValueMember = Nothing Then
                    Link.Tag = Row.Item(Me._ValueMember)
                End If
                Link.Top = Top
                Link.Left = Me._PaddingLeft
                Link.Height = Me._LineHeight
                Width = Me.Panel.Width - Me._PaddingLeft - Me._PaddingRight - 15
                If Width > 50 Then
                    Link.Width = Width
                Else
                    Link.Width = 50
                End If
                Link.TextAlign = ContentAlignment.TopLeft
                Link.ForeColor = Me._LinkColor
                Top += Me._LineHeight
                AddHandler Link.Click, AddressOf Clicked
                AddHandler Link.MouseHover, AddressOf Over
                AddHandler Link.MouseLeave, AddressOf Out
            Next
        End If
    End Sub

    Private Sub Over(ByVal sender As Object, ByVal e As EventArgs)
        Dim Link As LinkLabel = CType(sender, LinkLabel)
        Link.ForeColor = Me._HoverColor
    End Sub

    Private Sub Out(ByVal sender As Object, ByVal e As EventArgs)
        Dim Link As LinkLabel = CType(sender, LinkLabel)
        Link.ForeColor = Me._LinkColor
    End Sub

    Private Sub Clicked(ByVal sender As Object, ByVal e As EventArgs)
        RaiseEvent LinkClicked(sender.Tag)
    End Sub

    Private Sub _Table_RowChanging(ByVal sender As Object, ByVal e As System.Data.DataRowChangeEventArgs) Handles _Table.RowChanging
        Me.RedrawLinks()
    End Sub

    Private Sub _Table_RowDeleted(ByVal sender As Object, ByVal e As System.Data.DataRowChangeEventArgs) Handles _Table.RowDeleted
        Me.RedrawLinks()
    End Sub

    Private Sub Panel_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Panel.Resize
        'Me.RedrawLinks()
    End Sub


    Private Sub Panel_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel.Paint

    End Sub
End Class
