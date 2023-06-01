<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ArrowButtonPrev
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ArrowButtonPrev))
        Me.Button = New Infragistics.Win.Misc.UltraButton
        Me.SuspendLayout()
        '
        'Button
        '
        Appearance1.BackColor = System.Drawing.Color.DarkGray
        Appearance1.BackColor2 = System.Drawing.Color.Black
        Appearance1.ForeColor = System.Drawing.Color.White
        Me.Button.Appearance = Appearance1
        Me.Button.ButtonStyle = Infragistics.Win.UIElementButtonStyle.VisualStudio2005Button
        Me.Button.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button.Location = New System.Drawing.Point(0, 0)
        Me.Button.Name = "Button"
        Me.Button.ShapeImage = CType(resources.GetObject("Button.ShapeImage"), System.Drawing.Image)
        Me.Button.Size = New System.Drawing.Size(120, 43)
        Me.Button.TabIndex = 1
        Me.Button.Text = "Prev Page"
        '
        'ArrowButtonPrev
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.Button)
        Me.Name = "ArrowButtonPrev"
        Me.Size = New System.Drawing.Size(120, 43)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button As Infragistics.Win.Misc.UltraButton

End Class
