Public Class GravityInput
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
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
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents lblCaption As System.Windows.Forms.Label
    Friend WithEvents txtMaskedInput As System.Windows.Forms.TextBox
    Friend WithEvents txtInput As System.Windows.Forms.TextBox
    Friend WithEvents ubtnOK As Infragistics.Win.Misc.UltraButton
    Friend WithEvents ubtnCancel As Infragistics.Win.Misc.UltraButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(GravityInput))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.lblCaption = New System.Windows.Forms.Label
        Me.txtMaskedInput = New System.Windows.Forms.TextBox
        Me.txtInput = New System.Windows.Forms.TextBox
        Me.ubtnOK = New Infragistics.Win.Misc.UltraButton
        Me.ubtnCancel = New Infragistics.Win.Misc.UltraButton
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(16, 32)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(72, 72)
        Me.PictureBox1.TabIndex = 10
        Me.PictureBox1.TabStop = False
        '
        'lblCaption
        '
        Me.lblCaption.Location = New System.Drawing.Point(104, 32)
        Me.lblCaption.Name = "lblCaption"
        Me.lblCaption.Size = New System.Drawing.Size(264, 32)
        Me.lblCaption.TabIndex = 8
        '
        'txtMaskedInput
        '
        Me.txtMaskedInput.Location = New System.Drawing.Point(96, 72)
        Me.txtMaskedInput.Name = "txtMaskedInput"
        Me.txtMaskedInput.Size = New System.Drawing.Size(272, 20)
        Me.txtMaskedInput.TabIndex = 13
        '
        'txtInput
        '
        Me.txtInput.Location = New System.Drawing.Point(96, 72)
        Me.txtInput.Name = "txtInput"
        Me.txtInput.Size = New System.Drawing.Size(272, 20)
        Me.txtInput.TabIndex = 14
        '
        'ubtnOK
        '
        Me.ubtnOK.Location = New System.Drawing.Point(280, 112)
        Me.ubtnOK.Name = "ubtnOK"
        Me.ubtnOK.Size = New System.Drawing.Size(88, 24)
        Me.ubtnOK.TabIndex = 15
        Me.ubtnOK.Text = "O&K"
        '
        'ubtnCancel
        '
        Me.ubtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ubtnCancel.Location = New System.Drawing.Point(184, 112)
        Me.ubtnCancel.Name = "ubtnCancel"
        Me.ubtnCancel.Size = New System.Drawing.Size(88, 24)
        Me.ubtnCancel.TabIndex = 16
        Me.ubtnCancel.Text = "&Cancel"
        '
        'GravityInput
        '
        Me.AcceptButton = Me.ubtnOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.ubtnCancel
        Me.ClientSize = New System.Drawing.Size(376, 154)
        Me.ControlBox = False
        Me.Controls.Add(Me.ubtnCancel)
        Me.Controls.Add(Me.ubtnOK)
        Me.Controls.Add(Me.txtInput)
        Me.Controls.Add(Me.txtMaskedInput)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.lblCaption)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "GravityInput"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Gravity Input"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public ButtonValue As Gravity.Response = Gravity.Response.OK

    Private Sub GravityInput_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txtInput.Focus()
    End Sub

    Private Sub GravityInput_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Me.txtInput.Focus()
    End Sub

    Private Sub ubtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnOK.Click
        Me.ButtonValue = Gravity.Response.OK
        Me.Close()
    End Sub

    Private Sub ubtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnCancel.Click
        Me.ButtonValue = Gravity.Response.Cancel
        Me.Close()
    End Sub

End Class
