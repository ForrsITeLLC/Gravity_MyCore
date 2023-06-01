Public Class AskWindow
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ubtnNo As Infragistics.Win.Misc.UltraButton
    Friend WithEvents ubtnYes As Infragistics.Win.Misc.UltraButton
    Friend WithEvents ubtnCancel As Infragistics.Win.Misc.UltraButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AskWindow))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.lblCaption = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.ubtnNo = New Infragistics.Win.Misc.UltraButton
        Me.ubtnYes = New Infragistics.Win.Misc.UltraButton
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
        Me.PictureBox1.TabIndex = 6
        Me.PictureBox1.TabStop = False
        '
        'lblCaption
        '
        Me.lblCaption.Location = New System.Drawing.Point(96, 32)
        Me.lblCaption.Name = "lblCaption"
        Me.lblCaption.Size = New System.Drawing.Size(264, 48)
        Me.lblCaption.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(104, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(264, 48)
        Me.Label1.TabIndex = 4
        '
        'ubtnNo
        '
        Me.ubtnNo.Location = New System.Drawing.Point(280, 96)
        Me.ubtnNo.Name = "ubtnNo"
        Me.ubtnNo.Size = New System.Drawing.Size(75, 23)
        Me.ubtnNo.TabIndex = 9
        Me.ubtnNo.Text = "No"
        '
        'ubtnYes
        '
        Me.ubtnYes.Location = New System.Drawing.Point(200, 96)
        Me.ubtnYes.Name = "ubtnYes"
        Me.ubtnYes.Size = New System.Drawing.Size(75, 23)
        Me.ubtnYes.TabIndex = 10
        Me.ubtnYes.Text = "Yes"
        '
        'ubtnCancel
        '
        Me.ubtnCancel.Location = New System.Drawing.Point(120, 96)
        Me.ubtnCancel.Name = "ubtnCancel"
        Me.ubtnCancel.Size = New System.Drawing.Size(75, 23)
        Me.ubtnCancel.TabIndex = 11
        Me.ubtnCancel.Text = "Cancel"
        '
        'AskWindow
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(376, 146)
        Me.ControlBox = False
        Me.Controls.Add(Me.ubtnCancel)
        Me.Controls.Add(Me.ubtnYes)
        Me.Controls.Add(Me.ubtnNo)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.lblCaption)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AskWindow"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Gravity Question"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public ButtonValue As Gravity.Response = Gravity.Response.Yes

    Private Sub btnNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnNo.Click
        ButtonValue = Gravity.Response.No
        Me.Close()
    End Sub

    Private Sub btnYes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnYes.Click
        ButtonValue = Gravity.Response.Yes
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnCancel.Click
        ButtonValue = Gravity.Response.Cancel
        Me.Close()
    End Sub

End Class