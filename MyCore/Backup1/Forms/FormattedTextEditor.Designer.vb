<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormattedTextEditor
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormattedTextEditor))
        Me.uftxtBody = New Infragistics.Win.FormattedLinkLabel.UltraFormattedTextEditor
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.ToolStrip2 = New System.Windows.Forms.ToolStrip
        Me.tbtnLink = New System.Windows.Forms.ToolStripButton
        Me.tbtnBold = New System.Windows.Forms.ToolStripButton
        Me.btnImg = New System.Windows.Forms.ToolStripButton
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Panel1.SuspendLayout()
        Me.ToolStrip2.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'uftxtBody
        '
        Me.uftxtBody.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.uftxtBody.Dock = System.Windows.Forms.DockStyle.Fill
        Me.uftxtBody.Location = New System.Drawing.Point(0, 25)
        Me.uftxtBody.Name = "uftxtBody"
        Me.uftxtBody.Size = New System.Drawing.Size(920, 472)
        Me.uftxtBody.TabIndex = 13
        Me.uftxtBody.Value = ""
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 497)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(920, 43)
        Me.Panel1.TabIndex = 14
        '
        'btnCancel
        '
        Me.btnCancel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnCancel.Location = New System.Drawing.Point(5, 5)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(116, 33)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnOK.Location = New System.Drawing.Point(5, 5)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(113, 33)
        Me.btnOK.TabIndex = 0
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'ToolStrip2
        '
        Me.ToolStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tbtnLink, Me.tbtnBold, Me.btnImg})
        Me.ToolStrip2.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip2.Name = "ToolStrip2"
        Me.ToolStrip2.Size = New System.Drawing.Size(920, 25)
        Me.ToolStrip2.TabIndex = 15
        Me.ToolStrip2.Text = "ToolStrip2"
        '
        'tbtnLink
        '
        Me.tbtnLink.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tbtnLink.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Underline)
        Me.tbtnLink.ForeColor = System.Drawing.Color.Blue
        Me.tbtnLink.Image = CType(resources.GetObject("tbtnLink.Image"), System.Drawing.Image)
        Me.tbtnLink.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tbtnLink.Name = "tbtnLink"
        Me.tbtnLink.Size = New System.Drawing.Size(33, 22)
        Me.tbtnLink.Text = "Link"
        '
        'tbtnBold
        '
        Me.tbtnBold.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.tbtnBold.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.tbtnBold.Image = CType(resources.GetObject("tbtnBold.Image"), System.Drawing.Image)
        Me.tbtnBold.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tbtnBold.Name = "tbtnBold"
        Me.tbtnBold.Size = New System.Drawing.Size(67, 22)
        Me.tbtnBold.Text = "Font Style"
        '
        'btnImg
        '
        Me.btnImg.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnImg.Image = CType(resources.GetObject("btnImg.Image"), System.Drawing.Image)
        Me.btnImg.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnImg.Name = "btnImg"
        Me.btnImg.Size = New System.Drawing.Size(23, 22)
        Me.btnImg.Text = "ToolStripButton1"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnOK)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel2.Location = New System.Drawing.Point(797, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Padding = New System.Windows.Forms.Padding(5)
        Me.Panel2.Size = New System.Drawing.Size(123, 43)
        Me.Panel2.TabIndex = 2
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.btnCancel)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel3.Location = New System.Drawing.Point(671, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Padding = New System.Windows.Forms.Padding(5)
        Me.Panel3.Size = New System.Drawing.Size(126, 43)
        Me.Panel3.TabIndex = 3
        '
        'FormattedTextEditor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(920, 540)
        Me.Controls.Add(Me.uftxtBody)
        Me.Controls.Add(Me.ToolStrip2)
        Me.Controls.Add(Me.Panel1)
        Me.MinimizeBox = False
        Me.Name = "FormattedTextEditor"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Formatted Text Editor"
        Me.Panel1.ResumeLayout(False)
        Me.ToolStrip2.ResumeLayout(False)
        Me.ToolStrip2.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents uftxtBody As Infragistics.Win.FormattedLinkLabel.UltraFormattedTextEditor
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents ToolStrip2 As System.Windows.Forms.ToolStrip
    Friend WithEvents tbtnLink As System.Windows.Forms.ToolStripButton
    Friend WithEvents tbtnBold As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnImg As System.Windows.Forms.ToolStripButton
End Class
