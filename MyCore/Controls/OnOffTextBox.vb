Public Class OnOffTextBox
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
    Friend WithEvents Checkbox As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents TextBox As Infragistics.Win.UltraWinEditors.UltraTextEditor
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Checkbox = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.TextBox = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        CType(Me.TextBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Checkbox
        '
        Me.Checkbox.Dock = System.Windows.Forms.DockStyle.Left
        Me.Checkbox.Location = New System.Drawing.Point(0, 0)
        Me.Checkbox.Name = "Checkbox"
        Me.Checkbox.Size = New System.Drawing.Size(16, 24)
        Me.Checkbox.TabIndex = 0
        '
        'TextBox
        '
        Me.TextBox.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextBox.Location = New System.Drawing.Point(16, 0)
        Me.TextBox.Name = "TextBox"
        Me.TextBox.ReadOnly = True
        Me.TextBox.Size = New System.Drawing.Size(156, 21)
        Me.TextBox.TabIndex = 1
        '
        'OnOffTextBox
        '
        Me.Controls.Add(Me.TextBox)
        Me.Controls.Add(Me.Checkbox)
        Me.Name = "OnOffTextBox"
        Me.Size = New System.Drawing.Size(172, 24)
        CType(Me.TextBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Event CheckChanged()
    Public Shadows Event TextChanged()

    Public Property Checked() As Boolean
        Get
            Return Me.Checkbox.Checked
        End Get
        Set(ByVal Value As Boolean)
            Me.Checkbox.Checked = Value
        End Set
    End Property

    Public Overrides Property Text() As String
        Get
            Return Me.TextBox.Text
        End Get
        Set(ByVal Value As String)
            Me.TextBox.Text = Value
        End Set
    End Property

    Private Sub Checkbox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Checkbox.CheckedChanged
        If Me.Checkbox.Checked Then
            Me.TextBox.ReadOnly = False
        Else
            Me.TextBox.ReadOnly = True
        End If
        RaiseEvent CheckChanged()
    End Sub

    Private Sub TextBox_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox.ValueChanged
        RaiseEvent TextChanged()
    End Sub

End Class
