Public Class NeverEndingProgress
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
    Friend WithEvents OsProgressBar1 As OSProgress.OSProgressBar
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents lblClock As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(NeverEndingProgress))
        Me.OsProgressBar1 = New OSProgress.OSProgressBar
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.lblClock = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'OsProgressBar1
        '
        Me.OsProgressBar1.AutoProgress = True
        Me.OsProgressBar1.AutoProgressSpeed = CType(200, Byte)
        Me.OsProgressBar1.ForeColor = System.Drawing.Color.DarkGray
        Me.OsProgressBar1.IndicatorColor = System.Drawing.Color.MediumSeaGreen
        Me.OsProgressBar1.Location = New System.Drawing.Point(80, 32)
        Me.OsProgressBar1.Name = "OsProgressBar1"
        Me.OsProgressBar1.Position = 8
        Me.OsProgressBar1.ProgressBoxStyle = OSProgress.OSProgressBar.OSProgressBoxStyleConstants.osBOXAROUND
        Me.OsProgressBar1.Size = New System.Drawing.Size(208, 16)
        Me.OsProgressBar1.TabIndex = 0
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(288, 8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(64, 64)
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 8)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(64, 64)
        Me.PictureBox2.TabIndex = 2
        Me.PictureBox2.TabStop = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(80, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(200, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Connecting to Server..."
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(128, 72)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(88, 24)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "&Cancel"
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'lblClock
        '
        Me.lblClock.Location = New System.Drawing.Point(296, 80)
        Me.lblClock.Name = "lblClock"
        Me.lblClock.Size = New System.Drawing.Size(24, 16)
        Me.lblClock.TabIndex = 5
        Me.lblClock.Text = "30"
        Me.lblClock.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(248, 80)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 16)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Timeout"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(320, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(24, 16)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "sec"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'NeverEndingProgress
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(352, 114)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblClock)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.OsProgressBar1)
        Me.Name = "NeverEndingProgress"
        Me.ShowInTaskbar = False
        Me.Text = "Progress"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Canceled As Boolean = False
    Public ReachedTimeout As Boolean = False
    Public Loaded As Boolean = False
    Public Timeout As Integer = 30

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Canceled = True
        Me.Close()
    End Sub

    Private Sub NeverEndingProgress_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Me.Loaded = True
        Me.Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If Me.Timeout <= 0 Then
            Me.Timer1.Stop()
            Me.ReachedTimeout = True
            Me.Close()
        Else
            Me.Timeout -= 1
            Me.lblClock.Text = Me.Timeout
        End If
    End Sub

End Class
