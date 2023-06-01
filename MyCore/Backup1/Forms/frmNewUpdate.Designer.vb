<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNewUpdate
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNewUpdate))
        Me.lblTitle = New System.Windows.Forms.Label
        Me.txtDescription = New System.Windows.Forms.TextBox
        Me.lblDate = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnNow = New System.Windows.Forms.Button
        Me.btnLater = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblPriority = New System.Windows.Forms.Label
        Me.picPriority = New System.Windows.Forms.PictureBox
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        CType(Me.picPriority, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.Location = New System.Drawing.Point(12, 6)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(228, 17)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "Bug Fixes and Feature Update"
        '
        'txtDescription
        '
        Me.txtDescription.BackColor = System.Drawing.Color.White
        Me.txtDescription.Location = New System.Drawing.Point(15, 52)
        Me.txtDescription.Multiline = True
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.ReadOnly = True
        Me.txtDescription.Size = New System.Drawing.Size(346, 102)
        Me.txtDescription.TabIndex = 1
        '
        'lblDate
        '
        Me.lblDate.AutoSize = True
        Me.lblDate.Location = New System.Drawing.Point(65, 27)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(65, 13)
        Me.lblDate.TabIndex = 2
        Me.lblDate.Text = "10/04/2006"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Released:"
        '
        'btnNow
        '
        Me.btnNow.Location = New System.Drawing.Point(260, 160)
        Me.btnNow.Name = "btnNow"
        Me.btnNow.Size = New System.Drawing.Size(101, 23)
        Me.btnNow.TabIndex = 4
        Me.btnNow.Text = "Install Now"
        Me.btnNow.UseVisualStyleBackColor = True
        '
        'btnLater
        '
        Me.btnLater.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnLater.Location = New System.Drawing.Point(153, 160)
        Me.btnLater.Name = "btnLater"
        Me.btnLater.Size = New System.Drawing.Size(101, 23)
        Me.btnLater.TabIndex = 5
        Me.btnLater.Text = "Remind Me Later"
        Me.btnLater.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(161, 27)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Priority:"
        '
        'lblPriority
        '
        Me.lblPriority.AutoSize = True
        Me.lblPriority.Location = New System.Drawing.Point(208, 27)
        Me.lblPriority.Name = "lblPriority"
        Me.lblPriority.Size = New System.Drawing.Size(40, 13)
        Me.lblPriority.TabIndex = 7
        Me.lblPriority.Text = "Normal"
        '
        'picPriority
        '
        Me.picPriority.Image = CType(resources.GetObject("picPriority.Image"), System.Drawing.Image)
        Me.picPriority.Location = New System.Drawing.Point(315, 1)
        Me.picPriority.Name = "picPriority"
        Me.picPriority.Size = New System.Drawing.Size(48, 50)
        Me.picPriority.TabIndex = 8
        Me.picPriority.TabStop = False
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "Flag_1.gif")
        Me.ImageList1.Images.SetKeyName(1, "Flag_3.gif")
        '
        'frmNewUpdate
        '
        Me.AcceptButton = Me.btnNow
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.CancelButton = Me.btnLater
        Me.ClientSize = New System.Drawing.Size(373, 197)
        Me.ControlBox = False
        Me.Controls.Add(Me.picPriority)
        Me.Controls.Add(Me.lblPriority)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnLater)
        Me.Controls.Add(Me.btnNow)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblDate)
        Me.Controls.Add(Me.txtDescription)
        Me.Controls.Add(Me.lblTitle)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmNewUpdate"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Gravity CRM Update Available"
        CType(Me.picPriority, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnNow As System.Windows.Forms.Button
    Friend WithEvents btnLater As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblPriority As System.Windows.Forms.Label
    Friend WithEvents picPriority As System.Windows.Forms.PictureBox
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
End Class
