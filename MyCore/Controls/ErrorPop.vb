Public Class ErrorPop
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
    Friend WithEvents txtDetails As System.Windows.Forms.TextBox
    Friend WithEvents uchkNotify As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents uchkLog As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents ubtnCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents ubtnContinue As Infragistics.Win.Misc.UltraButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(ErrorPop))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.lblCaption = New System.Windows.Forms.Label
        Me.txtDetails = New System.Windows.Forms.TextBox
        Me.uchkNotify = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.uchkLog = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.ubtnCancel = New Infragistics.Win.Misc.UltraButton
        Me.ubtnContinue = New Infragistics.Win.Misc.UltraButton
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(16, 32)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(72, 72)
        Me.PictureBox1.TabIndex = 4
        Me.PictureBox1.TabStop = False
        '
        'lblCaption
        '
        Me.lblCaption.Location = New System.Drawing.Point(96, 32)
        Me.lblCaption.Name = "lblCaption"
        Me.lblCaption.Size = New System.Drawing.Size(264, 40)
        Me.lblCaption.TabIndex = 5
        '
        'txtDetails
        '
        Me.txtDetails.BackColor = System.Drawing.Color.White
        Me.txtDetails.Location = New System.Drawing.Point(96, 80)
        Me.txtDetails.Multiline = True
        Me.txtDetails.Name = "txtDetails"
        Me.txtDetails.ReadOnly = True
        Me.txtDetails.Size = New System.Drawing.Size(264, 72)
        Me.txtDetails.TabIndex = 6
        Me.txtDetails.Text = ""
        '
        'uchkNotify
        '
        Me.uchkNotify.Location = New System.Drawing.Point(8, 128)
        Me.uchkNotify.Name = "uchkNotify"
        Me.uchkNotify.Size = New System.Drawing.Size(88, 20)
        Me.uchkNotify.TabIndex = 11
        Me.uchkNotify.Text = "Notify Admin"
        '
        'uchkLog
        '
        Me.uchkLog.Location = New System.Drawing.Point(8, 104)
        Me.uchkLog.Name = "uchkLog"
        Me.uchkLog.Size = New System.Drawing.Size(72, 20)
        Me.uchkLog.TabIndex = 12
        Me.uchkLog.Text = "Log Error"
        '
        'ubtnCancel
        '
        Me.ubtnCancel.Location = New System.Drawing.Point(160, 168)
        Me.ubtnCancel.Name = "ubtnCancel"
        Me.ubtnCancel.Size = New System.Drawing.Size(104, 23)
        Me.ubtnCancel.TabIndex = 13
        Me.ubtnCancel.Text = "Close Program"
        '
        'ubtnContinue
        '
        Me.ubtnContinue.Location = New System.Drawing.Point(272, 168)
        Me.ubtnContinue.Name = "ubtnContinue"
        Me.ubtnContinue.Size = New System.Drawing.Size(88, 23)
        Me.ubtnContinue.TabIndex = 14
        Me.ubtnContinue.Text = "Continue"
        '
        'ErrorPop
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(376, 202)
        Me.ControlBox = False
        Me.Controls.Add(Me.ubtnContinue)
        Me.Controls.Add(Me.ubtnCancel)
        Me.Controls.Add(Me.uchkLog)
        Me.Controls.Add(Me.uchkNotify)
        Me.Controls.Add(Me.txtDetails)
        Me.Controls.Add(Me.lblCaption)
        Me.Controls.Add(Me.PictureBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ErrorPop"
        Me.Text = "Gravity Error"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public ButtonValue As Gravity.Response = Gravity.Response.OK

    Private Sub btnContinue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnContinue.Click
        ButtonValue = Gravity.Response.OK
        Me.Close()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnCancel.Click
        ButtonValue = Gravity.Response.Quit
        Me.Close()
    End Sub

    Private Sub ErrorPop_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        Me.ubtnCancel.Focus()
    End Sub


End Class
