Public Class frmDateBox
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
    Friend WithEvents lblBetween As System.Windows.Forms.Label
    Friend WithEvents lblAnd As System.Windows.Forms.Label
    Friend WithEvents udtpStart As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents udtpEnd As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents ubtnCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents pnlBetween As System.Windows.Forms.Panel
    Friend WithEvents ubtnOK As Infragistics.Win.Misc.UltraButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDateBox))
        Dim DateButton1 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim DateButton2 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.lblCaption = New System.Windows.Forms.Label
        Me.udtpStart = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.lblBetween = New System.Windows.Forms.Label
        Me.lblAnd = New System.Windows.Forms.Label
        Me.udtpEnd = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.ubtnCancel = New Infragistics.Win.Misc.UltraButton
        Me.ubtnOK = New Infragistics.Win.Misc.UltraButton
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.pnlBetween = New System.Windows.Forms.Panel
        Me.Panel4 = New System.Windows.Forms.Panel
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.udtpStart, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.udtpEnd, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.pnlBetween.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(10, 10)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(71, 56)
        Me.PictureBox1.TabIndex = 10
        Me.PictureBox1.TabStop = False
        '
        'lblCaption
        '
        Me.lblCaption.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblCaption.Location = New System.Drawing.Point(0, 0)
        Me.lblCaption.Name = "lblCaption"
        Me.lblCaption.Size = New System.Drawing.Size(376, 47)
        Me.lblCaption.TabIndex = 8
        Me.lblCaption.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'udtpStart
        '
        Me.udtpStart.DateButtons.Add(DateButton1)
        Me.udtpStart.Location = New System.Drawing.Point(6, 9)
        Me.udtpStart.Name = "udtpStart"
        Me.udtpStart.NonAutoSizeHeight = 23
        Me.udtpStart.Size = New System.Drawing.Size(88, 21)
        Me.udtpStart.TabIndex = 12
        '
        'lblBetween
        '
        Me.lblBetween.AutoSize = True
        Me.lblBetween.Location = New System.Drawing.Point(11, 14)
        Me.lblBetween.Name = "lblBetween"
        Me.lblBetween.Size = New System.Drawing.Size(49, 13)
        Me.lblBetween.TabIndex = 13
        Me.lblBetween.Text = "Between"
        Me.lblBetween.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblAnd
        '
        Me.lblAnd.AutoSize = True
        Me.lblAnd.Location = New System.Drawing.Point(34, 42)
        Me.lblAnd.Name = "lblAnd"
        Me.lblAnd.Size = New System.Drawing.Size(25, 13)
        Me.lblAnd.TabIndex = 14
        Me.lblAnd.Text = "and"
        Me.lblAnd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'udtpEnd
        '
        Me.udtpEnd.DateButtons.Add(DateButton2)
        Me.udtpEnd.Location = New System.Drawing.Point(6, 39)
        Me.udtpEnd.Name = "udtpEnd"
        Me.udtpEnd.NonAutoSizeHeight = 23
        Me.udtpEnd.Size = New System.Drawing.Size(88, 21)
        Me.udtpEnd.TabIndex = 15
        '
        'ubtnCancel
        '
        Me.ubtnCancel.Location = New System.Drawing.Point(192, 120)
        Me.ubtnCancel.Name = "ubtnCancel"
        Me.ubtnCancel.Size = New System.Drawing.Size(75, 23)
        Me.ubtnCancel.TabIndex = 16
        Me.ubtnCancel.Text = "&Cancel"
        '
        'ubtnOK
        '
        Me.ubtnOK.Location = New System.Drawing.Point(280, 120)
        Me.ubtnOK.Name = "ubtnOK"
        Me.ubtnOK.Size = New System.Drawing.Size(75, 23)
        Me.ubtnOK.TabIndex = 17
        Me.ubtnOK.Text = "&OK"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel1.Location = New System.Drawing.Point(0, 47)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Padding = New System.Windows.Forms.Padding(10)
        Me.Panel1.Size = New System.Drawing.Size(91, 109)
        Me.Panel1.TabIndex = 18
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Panel4)
        Me.Panel2.Controls.Add(Me.pnlBetween)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(91, 47)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(285, 67)
        Me.Panel2.TabIndex = 19
        '
        'pnlBetween
        '
        Me.pnlBetween.Controls.Add(Me.lblBetween)
        Me.pnlBetween.Controls.Add(Me.lblAnd)
        Me.pnlBetween.Dock = System.Windows.Forms.DockStyle.Left
        Me.pnlBetween.Location = New System.Drawing.Point(0, 0)
        Me.pnlBetween.Name = "pnlBetween"
        Me.pnlBetween.Size = New System.Drawing.Size(66, 67)
        Me.pnlBetween.TabIndex = 16
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.udtpStart)
        Me.Panel4.Controls.Add(Me.udtpEnd)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel4.Location = New System.Drawing.Point(66, 0)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(219, 67)
        Me.Panel4.TabIndex = 17
        '
        'frmDateBox
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(376, 156)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ubtnOK)
        Me.Controls.Add(Me.ubtnCancel)
        Me.Controls.Add(Me.lblCaption)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDateBox"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Date Range"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.udtpStart, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.udtpEnd, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.pnlBetween.ResumeLayout(False)
        Me.pnlBetween.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public ButtonValue As Gravity.Response = Gravity.Response.OK

    Private Sub ubtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnCancel.Click
        ButtonValue = Gravity.Response.Cancel
        Me.Close()
    End Sub

    Private Sub ubtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnOK.Click
        ButtonValue = Gravity.Response.OK
        Me.Close()
    End Sub

    Private Sub frmDateBox_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
