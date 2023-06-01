Public Class LookupBox
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
    Friend WithEvents TextBox As Infragistics.Win.UltraWinEditors.UltraTextEditor
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim EditorButton1 As Infragistics.Win.UltraWinEditors.EditorButton = New Infragistics.Win.UltraWinEditors.EditorButton("link")
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(LookupBox))
        Dim EditorButton2 As Infragistics.Win.UltraWinEditors.EditorButton = New Infragistics.Win.UltraWinEditors.EditorButton("new")
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim EditorButton3 As Infragistics.Win.UltraWinEditors.EditorButton = New Infragistics.Win.UltraWinEditors.EditorButton("lookup")
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.TextBox = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        CType(Me.TextBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TextBox
        '
        Appearance1.BorderColor = System.Drawing.Color.LightSteelBlue
        Me.TextBox.Appearance = Appearance1
        Me.TextBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance2.BackColor = System.Drawing.Color.Transparent
        Appearance2.FontData.UnderlineAsString = "True"
        Appearance2.ForeColor = System.Drawing.Color.Blue
        Appearance2.Image = CType(resources.GetObject("Appearance2.Image"), Object)
        EditorButton1.Appearance = Appearance2
        EditorButton1.Key = "link"
        EditorButton1.Text = ""
        EditorButton1.Visible = False
        Appearance3.BackColor = System.Drawing.Color.Transparent
        Appearance3.FontData.UnderlineAsString = "True"
        Appearance3.ForeColor = System.Drawing.Color.Blue
        Appearance3.Image = CType(resources.GetObject("Appearance3.Image"), Object)
        EditorButton2.Appearance = Appearance3
        EditorButton2.Key = "new"
        EditorButton2.Text = ""
        EditorButton2.Visible = False
        Appearance4.Image = CType(resources.GetObject("Appearance4.Image"), Object)
        EditorButton3.Appearance = Appearance4
        EditorButton3.Key = "lookup"
        EditorButton3.Text = ""
        Me.TextBox.ButtonsRight.Add(EditorButton1)
        Me.TextBox.ButtonsRight.Add(EditorButton2)
        Me.TextBox.ButtonsRight.Add(EditorButton3)
        Me.TextBox.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox.Location = New System.Drawing.Point(0, 0)
        Me.TextBox.Name = "TextBox"
        Me.TextBox.Size = New System.Drawing.Size(276, 21)
        Me.TextBox.TabIndex = 0
        '
        'LookupBox
        '
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.TextBox)
        Me.Name = "LookupBox"
        Me.Size = New System.Drawing.Size(276, 21)
        CType(Me.TextBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Dim _resolved As Boolean = False
    Dim _value As String = ""
    Dim _readonly As Boolean = False
    Dim _required As Boolean = False
    Dim _showlink As Boolean = False

    Dim _normalcolor As System.Drawing.Color
    Dim _typingcolor As System.Drawing.Color
    Dim _hovercolor As System.Drawing.Color

    Public Event ButtonClick(ByVal sender As Object, ByVal SearchText As String)
    Public Shadows Event TextChanged(ByVal sender As Object, ByVal SearchText As String)
    Public Shadows Event ValueChanged(ByVal sender As Object, ByVal Value As String)
    Public Shadows Event Leave(ByVal sender As Object, ByVal SearchText As String)
    Public Shadows Event Enter(ByVal sender As Object, ByVal SearchText As String)
    Public Shadows Event DoTest(ByVal sender As Object, ByVal SearchText As String)
    Public Event OpenClicked(ByVal sender As Object, ByVal Value As String)
    Public Event NewClicked(ByVal sender As Object, ByVal Value As String)

    Public Property ShowLink() As Boolean
        Get
            Return Me._showlink
        End Get
        Set(ByVal value As Boolean)
            Me._showlink = value
            Me.LinksRefresh()
        End Set
    End Property

    Public Overrides Property Text() As String
        Get
            Return Me.TextBox.Text
        End Get
        Set(ByVal Value As String)
            Me.TextBox.Text = Value
            RaiseEvent DoTest(Me, Me.TextBox.Text)
        End Set
    End Property

    Public ReadOnly Property Value() As String
        Get
            Return Me._value
        End Get
    End Property

    Public Property TextReadOnly() As Boolean
        Get
            Return Me._readonly
        End Get
        Set(ByVal Value As Boolean)
            Me.TextBox.ReadOnly = Value
            Me._readonly = Value
        End Set
    End Property

    Public ReadOnly Property Resolved() As Boolean
        Get
            Return Me.LocalResolved
        End Get
    End Property

    Public Property ShowButton() As Boolean
        Get
            Return Me.TextBox.ButtonsRight.Item("lookup").Visible
        End Get
        Set(ByVal Value As Boolean)
            Me.TextBox.ButtonsRight.Item("lookup").Visible = Value
        End Set
    End Property

    Public Property StyleSetName() As String
        Get
            Return Me.TextBox.StyleSetName
        End Get
        Set(ByVal value As String)
            Me.TextBox.StyleSetName = value
        End Set
    End Property

    Private Property LocalResolved() As Boolean
        Get
            Return _resolved
        End Get
        Set(ByVal Value As Boolean)
            Me._resolved = Value
            Me.LinksRefresh()
            If Value = True Then
                Me.SetFontMatched()
            Else
                Me.SetFontNoMatch()
                Me._value = ""
            End If
        End Set
    End Property

    Public Property ShowInkButton() As Infragistics.Win.ShowInkButton
        Get
            Return Me.TextBox.ShowInkButton
        End Get
        Set(ByVal value As Infragistics.Win.ShowInkButton)
            Me.TextBox.ShowInkButton = value
        End Set
    End Property

    Private Sub LookupBox_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox_EditorButtonClick(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinEditors.EditorButtonEventArgs) Handles TextBox.EditorButtonClick
        If e.Button.Key = "lookup" Then
            RaiseEvent ButtonClick(Me, Me.TextBox.Text)
        ElseIf e.Button.Key = "link" Then
            If Me._resolved Then
                RaiseEvent OpenClicked(Me, Me.Value)
            End If
        ElseIf e.Button.Key = "new" Then
            RaiseEvent NewClicked(Me, Me.Value)
        End If
    End Sub

    Private Sub TextBox_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox.ValueChanged
        RaiseEvent TextChanged(Me, Me.TextBox.Text)
    End Sub

    Private Sub TextBox_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox.Leave
        Me.SuspendLayout()
        RaiseEvent Leave(Me, Me.TextBox.Text)
        Me.TextBox.Appearance.BackColor = Color.White
        If Not Me._readonly Then
            RaiseEvent DoTest(Me, Me.TextBox.Text)
        End If
        Me.ResumeLayout()
    End Sub

    Private Sub TextBox_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox.Enter
        Me.SuspendLayout()
        RaiseEvent Enter(Me, Me.TextBox.Text)
        Me.TextBox.Appearance.BackColor = Color.WhiteSmoke
        If Not Me._readonly Then
            Me.SetFontTyping()
            If Me._resolved Then
                Me.TextBox.Text = Value
            End If
        End If
        Me.ResumeLayout()
    End Sub

    Public Sub FoundMatch(ByVal Value As String, ByVal Text As String)
        If Value <> Me._value Or Not Me.Resolved Then
            Me._value = Value
            Me.TextBox.Text = Text
            Me.LocalResolved = True
            RaiseEvent ValueChanged(Me, Value)
        Else
            Me.TextBox.Text = Text
            Me.LocalResolved = True
        End If
    End Sub

    Public Sub NoMatch()
        Me._value = ""
        Me.LocalResolved = False
        RaiseEvent ValueChanged(Me, "")
    End Sub

    Public Sub Clear()
        If Me.Value.Length > 0 Then
            Me._value = ""
            Me.Text = ""
            Me.LocalResolved = False
            Me.SetFontTyping()
            RaiseEvent ValueChanged(Me, "")
        End If
    End Sub

    Public Sub Search()
        RaiseEvent DoTest(Me, Me.TextBox.Text)
    End Sub

    Private Sub TextBox_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox.KeyDown
        If Me._readonly Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox.KeyPress
        If Me._readonly Then
            e.Handled = True
        End If
    End Sub

    Private Sub SetFontNoMatch()
        Me.TextBox.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.False
        Me.TextBox.Appearance.FontData.Underline = Infragistics.Win.DefaultableBoolean.False
        Me.TextBox.Appearance.ForeColor = System.Drawing.Color.Red
        Me.TextBox.Appearance.FontData.Italic = Infragistics.Win.DefaultableBoolean.True
    End Sub

    Private Sub SetFontMatched()
        Me.TextBox.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
        Me.TextBox.Appearance.FontData.Underline = Infragistics.Win.DefaultableBoolean.True
        Me.TextBox.Appearance.FontData.Italic = Infragistics.Win.DefaultableBoolean.False
        Me.TextBox.Appearance.ForeColor = System.Drawing.Color.Black
    End Sub

    Private Sub SetFontTyping()
        Me.TextBox.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.False
        Me.TextBox.Appearance.FontData.Underline = Infragistics.Win.DefaultableBoolean.False
        Me.TextBox.Appearance.ForeColor = System.Drawing.Color.Black
        Me.TextBox.Appearance.FontData.Italic = Infragistics.Win.DefaultableBoolean.False
    End Sub

    Private Sub LinksRefresh()
        Me.TextBox.ButtonsRight.Item("link").Visible = False
        Me.TextBox.ButtonsRight.Item("new").Visible = False
        If Me.ShowLink Then
            If Me._resolved Then
                Me.TextBox.ButtonsRight.Item("link").Visible = True
            Else
                Me.TextBox.ButtonsRight.Item("new").Visible = True
            End If
        End If
    End Sub

End Class
