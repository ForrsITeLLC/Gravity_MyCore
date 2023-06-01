Public Class frmSelect
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
    Friend WithEvents ucboChoice As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents ubtnCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents ubtnOK As Infragistics.Win.Misc.UltraButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSelect))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.lblCaption = New System.Windows.Forms.Label
        Me.ucboChoice = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.ubtnCancel = New Infragistics.Win.Misc.UltraButton
        Me.ubtnOK = New Infragistics.Win.Misc.UltraButton
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ucboChoice, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(72, 64)
        Me.PictureBox1.TabIndex = 10
        Me.PictureBox1.TabStop = False
        '
        'lblCaption
        '
        Me.lblCaption.Location = New System.Drawing.Point(88, 16)
        Me.lblCaption.Name = "lblCaption"
        Me.lblCaption.Size = New System.Drawing.Size(264, 32)
        Me.lblCaption.TabIndex = 8
        '
        'ucboChoice
        '
        Me.ucboChoice.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.ucboChoice.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.[Default]
        Me.ucboChoice.Location = New System.Drawing.Point(88, 48)
        Me.ucboChoice.Name = "ucboChoice"
        Me.ucboChoice.Size = New System.Drawing.Size(264, 22)
        Me.ucboChoice.TabIndex = 12
        '
        'ubtnCancel
        '
        Me.ubtnCancel.Location = New System.Drawing.Point(168, 80)
        Me.ubtnCancel.Name = "ubtnCancel"
        Me.ubtnCancel.Size = New System.Drawing.Size(88, 23)
        Me.ubtnCancel.TabIndex = 13
        Me.ubtnCancel.Text = "Cancel"
        '
        'ubtnOK
        '
        Me.ubtnOK.Location = New System.Drawing.Point(264, 80)
        Me.ubtnOK.Name = "ubtnOK"
        Me.ubtnOK.Size = New System.Drawing.Size(88, 23)
        Me.ubtnOK.TabIndex = 14
        Me.ubtnOK.Text = "OK"
        '
        'frmSelect
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(362, 116)
        Me.ControlBox = False
        Me.Controls.Add(Me.ubtnOK)
        Me.Controls.Add(Me.ubtnCancel)
        Me.Controls.Add(Me.ucboChoice)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.lblCaption)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSelect"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Gravity Select"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ucboChoice, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public ButtonValue As Gravity.Response = Gravity.Response.OK

    Private Sub ucboSelect_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub frmSelect_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        Me.ucboChoice.Focus()
    End Sub

    Private Sub ubtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnOK.Click
        ButtonValue = Gravity.Response.OK
        Me.Close()
    End Sub

    Private Sub ubtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnCancel.Click
        ButtonValue = Gravity.Response.Cancel
        Me.Close()
    End Sub

End Class
