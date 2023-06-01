<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLogin
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.lnkContinue = New System.Windows.Forms.LinkLabel
        Me.lblLogin = New Infragistics.Win.Misc.UltraLabel
        Me.utxtUser = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.utxtPassword = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.uchkRemember = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        CType(Me.utxtUser, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.utxtPassword, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lnkContinue
        '
        Me.lnkContinue.AutoSize = True
        Me.lnkContinue.Location = New System.Drawing.Point(319, 248)
        Me.lnkContinue.Name = "lnkContinue"
        Me.lnkContinue.Size = New System.Drawing.Size(134, 13)
        Me.lnkContinue.TabIndex = 31
        Me.lnkContinue.TabStop = True
        Me.lnkContinue.Text = "Continue without logging in"
        '
        'lblLogin
        '
        Appearance1.BackColor = System.Drawing.Color.Transparent
        Me.lblLogin.Appearance = Appearance1
        Me.lblLogin.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lblLogin.Location = New System.Drawing.Point(481, 217)
        Me.lblLogin.Name = "lblLogin"
        Me.lblLogin.Size = New System.Drawing.Size(100, 23)
        Me.lblLogin.TabIndex = 30
        '
        'utxtUser
        '
        Me.utxtUser.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        Me.utxtUser.Location = New System.Drawing.Point(326, 126)
        Me.utxtUser.Name = "utxtUser"
        Me.utxtUser.Size = New System.Drawing.Size(230, 17)
        Me.utxtUser.TabIndex = 27
        '
        'utxtPassword
        '
        Me.utxtPassword.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        Me.utxtPassword.Location = New System.Drawing.Point(326, 179)
        Me.utxtPassword.Name = "utxtPassword"
        Me.utxtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.utxtPassword.Size = New System.Drawing.Size(230, 17)
        Me.utxtPassword.TabIndex = 28
        '
        'uchkRemember
        '
        Me.uchkRemember.BackColor = System.Drawing.Color.Transparent
        Me.uchkRemember.BackColorInternal = System.Drawing.Color.Transparent
        Me.uchkRemember.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Borderless
        Me.uchkRemember.Location = New System.Drawing.Point(322, 219)
        Me.uchkRemember.Name = "uchkRemember"
        Me.uchkRemember.Size = New System.Drawing.Size(120, 20)
        Me.uchkRemember.TabIndex = 29
        '
        'frmLogin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.MyCore.My.Resources.Resources.gravity_login
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.ClientSize = New System.Drawing.Size(600, 280)
        Me.ControlBox = False
        Me.Controls.Add(Me.lnkContinue)
        Me.Controls.Add(Me.lblLogin)
        Me.Controls.Add(Me.uchkRemember)
        Me.Controls.Add(Me.utxtUser)
        Me.Controls.Add(Me.utxtPassword)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLogin"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login"
        CType(Me.utxtUser, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.utxtPassword, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lnkContinue As System.Windows.Forms.LinkLabel
    Friend WithEvents lblLogin As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents utxtUser As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents utxtPassword As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents uchkRemember As Infragistics.Win.UltraWinEditors.UltraCheckEditor
End Class
