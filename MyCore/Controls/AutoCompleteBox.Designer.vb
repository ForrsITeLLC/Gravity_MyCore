<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AutoCompleteBox
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AutoCompleteBox))
        Me.TextBox = New System.Windows.Forms.TextBox
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnOpen = New System.Windows.Forms.Button
        Me.btnSearch = New System.Windows.Forms.Button
        Me.ToolTip2 = New System.Windows.Forms.ToolTip(Me.components)
        Me.pnlTextBox = New System.Windows.Forms.Panel
        Me.pnlTextBox.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox
        '
        Me.TextBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.TextBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.TextBox.BackColor = System.Drawing.Color.White
        Me.TextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox.Location = New System.Drawing.Point(0, 0)
        Me.TextBox.Name = "TextBox"
        Me.TextBox.Size = New System.Drawing.Size(156, 21)
        Me.TextBox.TabIndex = 0
        '
        'ToolTip1
        '
        Me.ToolTip1.AutomaticDelay = 200
        '
        'btnNew
        '
        Me.btnNew.BackColor = System.Drawing.Color.Gainsboro
        Me.btnNew.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnNew.Image = CType(resources.GetObject("btnNew.Image"), System.Drawing.Image)
        Me.btnNew.Location = New System.Drawing.Point(179, 0)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(23, 22)
        Me.btnNew.TabIndex = 4
        Me.btnNew.TabStop = False
        Me.ToolTip2.SetToolTip(Me.btnNew, "New")
        Me.btnNew.UseVisualStyleBackColor = False
        '
        'btnOpen
        '
        Me.btnOpen.BackColor = System.Drawing.Color.Gainsboro
        Me.btnOpen.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnOpen.Image = CType(resources.GetObject("btnOpen.Image"), System.Drawing.Image)
        Me.btnOpen.Location = New System.Drawing.Point(156, 0)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(23, 22)
        Me.btnOpen.TabIndex = 5
        Me.btnOpen.TabStop = False
        Me.ToolTip2.SetToolTip(Me.btnOpen, "Open")
        Me.btnOpen.UseVisualStyleBackColor = False
        '
        'btnSearch
        '
        Me.btnSearch.BackColor = System.Drawing.Color.Gainsboro
        Me.btnSearch.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnSearch.Image = CType(resources.GetObject("btnSearch.Image"), System.Drawing.Image)
        Me.btnSearch.Location = New System.Drawing.Point(202, 0)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(23, 22)
        Me.btnSearch.TabIndex = 6
        Me.btnSearch.TabStop = False
        Me.ToolTip2.SetToolTip(Me.btnSearch, "Search")
        Me.btnSearch.UseVisualStyleBackColor = False
        '
        'pnlTextBox
        '
        Me.pnlTextBox.BackColor = System.Drawing.Color.Gray
        Me.pnlTextBox.Controls.Add(Me.TextBox)
        Me.pnlTextBox.Location = New System.Drawing.Point(0, 0)
        Me.pnlTextBox.Name = "pnlTextBox"
        Me.pnlTextBox.Size = New System.Drawing.Size(156, 22)
        Me.pnlTextBox.TabIndex = 7
        '
        'AutoCompleteBox
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Red
        Me.Controls.Add(Me.pnlTextBox)
        Me.Controls.Add(Me.btnOpen)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnSearch)
        Me.Name = "AutoCompleteBox"
        Me.Size = New System.Drawing.Size(225, 22)
        Me.pnlTextBox.ResumeLayout(False)
        Me.pnlTextBox.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TextBox As System.Windows.Forms.TextBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnOpen As System.Windows.Forms.Button
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents ToolTip2 As System.Windows.Forms.ToolTip
    Friend WithEvents pnlTextBox As System.Windows.Forms.Panel

End Class
