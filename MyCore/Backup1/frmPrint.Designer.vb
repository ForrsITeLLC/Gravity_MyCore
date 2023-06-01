<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPrint
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
        Me.components = New System.ComponentModel.Container
        Me.Label1 = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.pnlOptions = New System.Windows.Forms.Panel
        Me.ubtnPrint = New Infragistics.Win.Misc.UltraButton
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.ubtnSuccess = New Infragistics.Win.Misc.UltraButton
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.ubtnFailed = New Infragistics.Win.Misc.UltraButton
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.pnlOptions.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label1.Location = New System.Drawing.Point(10, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(346, 48)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Preparing to print..."
        '
        'Timer1
        '
        '
        'WebBrowser1
        '
        Me.WebBrowser1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WebBrowser1.Location = New System.Drawing.Point(10, 10)
        Me.WebBrowser1.MinimumSize = New System.Drawing.Size(20, 20)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(346, 20)
        Me.WebBrowser1.TabIndex = 2
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.WebBrowser1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 100)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Padding = New System.Windows.Forms.Padding(10)
        Me.Panel1.Size = New System.Drawing.Size(366, 10)
        Me.Panel1.TabIndex = 3
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Padding = New System.Windows.Forms.Padding(10)
        Me.Panel2.Size = New System.Drawing.Size(366, 68)
        Me.Panel2.TabIndex = 4
        '
        'pnlOptions
        '
        Me.pnlOptions.Controls.Add(Me.ubtnPrint)
        Me.pnlOptions.Controls.Add(Me.Panel4)
        Me.pnlOptions.Controls.Add(Me.ubtnSuccess)
        Me.pnlOptions.Controls.Add(Me.Panel6)
        Me.pnlOptions.Controls.Add(Me.ubtnFailed)
        Me.pnlOptions.Controls.Add(Me.Panel5)
        Me.pnlOptions.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlOptions.Location = New System.Drawing.Point(0, 68)
        Me.pnlOptions.Name = "pnlOptions"
        Me.pnlOptions.Padding = New System.Windows.Forms.Padding(3)
        Me.pnlOptions.Size = New System.Drawing.Size(366, 32)
        Me.pnlOptions.TabIndex = 5
        Me.pnlOptions.Visible = False
        '
        'ubtnPrint
        '
        Me.ubtnPrint.ButtonStyle = Infragistics.Win.UIElementButtonStyle.VisualStudio2005Button
        Me.ubtnPrint.Dock = System.Windows.Forms.DockStyle.Right
        Me.ubtnPrint.Location = New System.Drawing.Point(39, 3)
        Me.ubtnPrint.Name = "ubtnPrint"
        Me.ubtnPrint.Size = New System.Drawing.Size(96, 26)
        Me.ubtnPrint.TabIndex = 18
        Me.ubtnPrint.Text = "Print Again"
        '
        'Panel4
        '
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel4.Location = New System.Drawing.Point(135, 3)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(12, 26)
        Me.Panel4.TabIndex = 17
        '
        'ubtnSuccess
        '
        Me.ubtnSuccess.ButtonStyle = Infragistics.Win.UIElementButtonStyle.VisualStudio2005Button
        Me.ubtnSuccess.Dock = System.Windows.Forms.DockStyle.Right
        Me.ubtnSuccess.Location = New System.Drawing.Point(147, 3)
        Me.ubtnSuccess.Name = "ubtnSuccess"
        Me.ubtnSuccess.Size = New System.Drawing.Size(96, 26)
        Me.ubtnSuccess.TabIndex = 16
        Me.ubtnSuccess.Text = "Yes"
        '
        'Panel6
        '
        Me.Panel6.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel6.Location = New System.Drawing.Point(243, 3)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(12, 26)
        Me.Panel6.TabIndex = 21
        '
        'ubtnFailed
        '
        Me.ubtnFailed.ButtonStyle = Infragistics.Win.UIElementButtonStyle.VisualStudio2005Button
        Me.ubtnFailed.Dock = System.Windows.Forms.DockStyle.Right
        Me.ubtnFailed.Location = New System.Drawing.Point(255, 3)
        Me.ubtnFailed.Name = "ubtnFailed"
        Me.ubtnFailed.Size = New System.Drawing.Size(96, 26)
        Me.ubtnFailed.TabIndex = 20
        Me.ubtnFailed.Text = "No"
        '
        'Panel5
        '
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel5.Location = New System.Drawing.Point(351, 3)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(12, 26)
        Me.Panel5.TabIndex = 19
        '
        'frmPrint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(366, 110)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.pnlOptions)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmPrint"
        Me.ShowIcon = False
        Me.Text = "Printing"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.pnlOptions.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents WebBrowser1 As System.Windows.Forms.WebBrowser
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents pnlOptions As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents ubtnSuccess As Infragistics.Win.Misc.UltraButton
    Friend WithEvents ubtnPrint As Infragistics.Win.Misc.UltraButton
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents ubtnFailed As Infragistics.Win.Misc.UltraButton
End Class
