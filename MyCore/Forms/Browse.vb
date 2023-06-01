Imports System.Windows.Forms

Public Class Browse
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
    Friend WithEvents pnlBottom As System.Windows.Forms.Panel
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents ubtnCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents lblFolders As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents pnlButtons As System.Windows.Forms.Panel
    Friend WithEvents FolderContext As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents mnuAddFolder As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuRenameFolder As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents pnlNew As System.Windows.Forms.Panel
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents lvFolders As System.Windows.Forms.ListView
    Friend WithEvents lblTemplates As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents lvDocs As System.Windows.Forms.ListView
    Friend WithEvents DocContext As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents mnuRenameDoc As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuDeleteTemplate As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImageList2 As System.Windows.Forms.ImageList
    Friend WithEvents ubtnOK As Infragistics.Win.Misc.UltraButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Browse))
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.pnlBottom = New System.Windows.Forms.Panel
        Me.pnlNew = New System.Windows.Forms.Panel
        Me.lblName = New System.Windows.Forms.Label
        Me.txtName = New System.Windows.Forms.TextBox
        Me.pnlButtons = New System.Windows.Forms.Panel
        Me.ubtnCancel = New Infragistics.Win.Misc.UltraButton
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.ubtnOK = New Infragistics.Win.Misc.UltraButton
        Me.FolderContext = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnuAddFolder = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRenameFolder = New System.Windows.Forms.ToolStripMenuItem
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.lblFolders = New Infragistics.Win.Misc.UltraLabel
        Me.lvFolders = New System.Windows.Forms.ListView
        Me.lblTemplates = New Infragistics.Win.Misc.UltraLabel
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.lvDocs = New System.Windows.Forms.ListView
        Me.DocContext = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnuRenameDoc = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuDeleteTemplate = New System.Windows.Forms.ToolStripMenuItem
        Me.ImageList2 = New System.Windows.Forms.ImageList(Me.components)
        Me.pnlBottom.SuspendLayout()
        Me.pnlNew.SuspendLayout()
        Me.pnlButtons.SuspendLayout()
        Me.FolderContext.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.DocContext.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlBottom
        '
        Me.pnlBottom.BackColor = System.Drawing.Color.WhiteSmoke
        Me.pnlBottom.Controls.Add(Me.pnlNew)
        Me.pnlBottom.Controls.Add(Me.pnlButtons)
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.Location = New System.Drawing.Point(0, 362)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(749, 60)
        Me.pnlBottom.TabIndex = 2
        '
        'pnlNew
        '
        Me.pnlNew.Controls.Add(Me.lblName)
        Me.pnlNew.Controls.Add(Me.txtName)
        Me.pnlNew.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlNew.Location = New System.Drawing.Point(0, 0)
        Me.pnlNew.Name = "pnlNew"
        Me.pnlNew.Padding = New System.Windows.Forms.Padding(5)
        Me.pnlNew.Size = New System.Drawing.Size(559, 60)
        Me.pnlNew.TabIndex = 8
        '
        'lblName
        '
        Me.lblName.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblName.Location = New System.Drawing.Point(5, 5)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(549, 30)
        Me.lblName.TabIndex = 2
        Me.lblName.Text = "Name:"
        Me.lblName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtName
        '
        Me.txtName.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.txtName.Location = New System.Drawing.Point(5, 35)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(549, 20)
        Me.txtName.TabIndex = 0
        '
        'pnlButtons
        '
        Me.pnlButtons.Controls.Add(Me.ubtnCancel)
        Me.pnlButtons.Controls.Add(Me.Panel2)
        Me.pnlButtons.Controls.Add(Me.ubtnOK)
        Me.pnlButtons.Dock = System.Windows.Forms.DockStyle.Right
        Me.pnlButtons.Location = New System.Drawing.Point(559, 0)
        Me.pnlButtons.Name = "pnlButtons"
        Me.pnlButtons.Padding = New System.Windows.Forms.Padding(5)
        Me.pnlButtons.Size = New System.Drawing.Size(190, 60)
        Me.pnlButtons.TabIndex = 14
        '
        'ubtnCancel
        '
        Me.ubtnCancel.ButtonStyle = Infragistics.Win.UIElementButtonStyle.WindowsXPCommandButton
        Me.ubtnCancel.Dock = System.Windows.Forms.DockStyle.Right
        Me.ubtnCancel.Location = New System.Drawing.Point(25, 5)
        Me.ubtnCancel.Margin = New System.Windows.Forms.Padding(3, 3, 10, 3)
        Me.ubtnCancel.Name = "ubtnCancel"
        Me.ubtnCancel.Size = New System.Drawing.Size(75, 50)
        Me.ubtnCancel.TabIndex = 12
        Me.ubtnCancel.Text = "Cancel"
        '
        'Panel2
        '
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel2.Location = New System.Drawing.Point(100, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(10, 50)
        Me.Panel2.TabIndex = 14
        '
        'ubtnOK
        '
        Appearance1.BorderColor = System.Drawing.Color.Gray
        Me.ubtnOK.Appearance = Appearance1
        Me.ubtnOK.ButtonStyle = Infragistics.Win.UIElementButtonStyle.WindowsXPCommandButton
        Me.ubtnOK.Dock = System.Windows.Forms.DockStyle.Right
        Me.ubtnOK.Location = New System.Drawing.Point(110, 5)
        Me.ubtnOK.Margin = New System.Windows.Forms.Padding(10, 3, 3, 3)
        Me.ubtnOK.Name = "ubtnOK"
        Me.ubtnOK.Size = New System.Drawing.Size(75, 50)
        Me.ubtnOK.TabIndex = 13
        Me.ubtnOK.Text = "OK"
        '
        'FolderContext
        '
        Me.FolderContext.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAddFolder, Me.mnuRenameFolder})
        Me.FolderContext.Name = "FolderContext"
        Me.FolderContext.Size = New System.Drawing.Size(147, 48)
        '
        'mnuAddFolder
        '
        Me.mnuAddFolder.Name = "mnuAddFolder"
        Me.mnuAddFolder.Size = New System.Drawing.Size(146, 22)
        Me.mnuAddFolder.Text = "Add Folder"
        '
        'mnuRenameFolder
        '
        Me.mnuRenameFolder.Name = "mnuRenameFolder"
        Me.mnuRenameFolder.Size = New System.Drawing.Size(146, 22)
        Me.mnuRenameFolder.Text = "Rename Folder"
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "")
        Me.ImageList1.Images.SetKeyName(1, "")
        Me.ImageList1.Images.SetKeyName(2, "")
        '
        'lblFolders
        '
        Appearance2.BackColor = System.Drawing.Color.WhiteSmoke
        Appearance2.TextVAlignAsString = "Middle"
        Me.lblFolders.Appearance = Appearance2
        Me.lblFolders.BorderStyleInner = Infragistics.Win.UIElementBorderStyle.Solid
        Me.lblFolders.BorderStyleOuter = Infragistics.Win.UIElementBorderStyle.None
        Me.lblFolders.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblFolders.Location = New System.Drawing.Point(5, 5)
        Me.lblFolders.Name = "lblFolders"
        Me.lblFolders.Size = New System.Drawing.Size(199, 23)
        Me.lblFolders.TabIndex = 5
        Me.lblFolders.Text = "Folders"
        '
        'lvFolders
        '
        Me.lvFolders.AllowDrop = True
        Me.lvFolders.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lvFolders.ContextMenuStrip = Me.FolderContext
        Me.lvFolders.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvFolders.LargeImageList = Me.ImageList1
        Me.lvFolders.Location = New System.Drawing.Point(5, 28)
        Me.lvFolders.MultiSelect = False
        Me.lvFolders.Name = "lvFolders"
        Me.lvFolders.Size = New System.Drawing.Size(199, 329)
        Me.lvFolders.SmallImageList = Me.ImageList1
        Me.lvFolders.TabIndex = 6
        Me.lvFolders.UseCompatibleStateImageBehavior = False
        Me.lvFolders.View = System.Windows.Forms.View.List
        '
        'lblTemplates
        '
        Appearance3.BackColor = System.Drawing.Color.WhiteSmoke
        Appearance3.TextVAlignAsString = "Middle"
        Me.lblTemplates.Appearance = Appearance3
        Me.lblTemplates.BorderStyleInner = Infragistics.Win.UIElementBorderStyle.Solid
        Me.lblTemplates.BorderStyleOuter = Infragistics.Win.UIElementBorderStyle.None
        Me.lblTemplates.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblTemplates.Location = New System.Drawing.Point(5, 5)
        Me.lblTemplates.Name = "lblTemplates"
        Me.lblTemplates.Size = New System.Drawing.Size(526, 23)
        Me.lblTemplates.TabIndex = 7
        Me.lblTemplates.Text = "Templates"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.lvFolders)
        Me.SplitContainer1.Panel1.Controls.Add(Me.lblFolders)
        Me.SplitContainer1.Panel1.Padding = New System.Windows.Forms.Padding(5)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.lvDocs)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblTemplates)
        Me.SplitContainer1.Panel2.Padding = New System.Windows.Forms.Padding(5)
        Me.SplitContainer1.Size = New System.Drawing.Size(749, 362)
        Me.SplitContainer1.SplitterDistance = 209
        Me.SplitContainer1.TabIndex = 8
        '
        'lvDocs
        '
        Me.lvDocs.AllowDrop = True
        Me.lvDocs.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lvDocs.ContextMenuStrip = Me.DocContext
        Me.lvDocs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvDocs.LargeImageList = Me.ImageList2
        Me.lvDocs.Location = New System.Drawing.Point(5, 28)
        Me.lvDocs.MultiSelect = False
        Me.lvDocs.Name = "lvDocs"
        Me.lvDocs.Size = New System.Drawing.Size(526, 329)
        Me.lvDocs.SmallImageList = Me.ImageList1
        Me.lvDocs.TabIndex = 8
        Me.lvDocs.UseCompatibleStateImageBehavior = False
        '
        'DocContext
        '
        Me.DocContext.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuRenameDoc, Me.mnuDeleteTemplate})
        Me.DocContext.Name = "DocContext"
        Me.DocContext.Size = New System.Drawing.Size(161, 48)
        '
        'mnuRenameDoc
        '
        Me.mnuRenameDoc.Name = "mnuRenameDoc"
        Me.mnuRenameDoc.Size = New System.Drawing.Size(160, 22)
        Me.mnuRenameDoc.Text = "Rename Template"
        '
        'mnuDeleteTemplate
        '
        Me.mnuDeleteTemplate.Name = "mnuDeleteTemplate"
        Me.mnuDeleteTemplate.Size = New System.Drawing.Size(160, 22)
        Me.mnuDeleteTemplate.Text = "Delete Template"
        '
        'ImageList2
        '
        Me.ImageList2.ImageStream = CType(resources.GetObject("ImageList2.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList2.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList2.Images.SetKeyName(0, "")
        Me.ImageList2.Images.SetKeyName(1, "")
        Me.ImageList2.Images.SetKeyName(2, "")
        '
        'Browse
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(749, 422)
        Me.ControlBox = False
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.pnlBottom)
        Me.Name = "Browse"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Browse"
        Me.pnlBottom.ResumeLayout(False)
        Me.pnlNew.ResumeLayout(False)
        Me.pnlNew.PerformLayout()
        Me.pnlButtons.ResumeLayout(False)
        Me.FolderContext.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.DocContext.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim _SelectedFolder As Integer = Nothing

    Public InitialFolder As Integer = 0
    Public Controller As MyCore.Gravity.ChooseTemplateDialog
    Public CurrentFolder As DataTable
    Public Filter As String = ""

    Dim DragTemplate As System.Windows.Forms.ListViewItem

    Public ReadOnly Property SelectedFolder() As Integer
        Get
            Return _SelectedFolder
        End Get
    End Property

    Public ReadOnly Property SelectedDocument() As Integer
        Get
            If Me.lvDocs.SelectedItems.Count > 0 Then
                Return Me.lvDocs.SelectedItems(0).Tag
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property SelectedTemplateHtml() As String
        Get
            If Me.lvDocs.SelectedItems.Count > 0 Then
                Return Me.Controller.Database.GetOne("SELECT html FROM template WHERE id=" & Me.lvDocs.SelectedItems(0).Tag)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Private Sub Browse_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.LoadFolders()
        ' If initial folder
        If Me.InitialFolder > 0 Then
            Me.OpenFolder(Me.InitialFolder)
        End If
    End Sub

    Public Sub LoadFolders()
        ' Clear any current
        Me.lvFolders.Items.Clear()
        ' Get Folders
        Dim Sql As String = "SELECT * FROM form_type ORDER BY name"
        Dim Table As DataTable = Me.Controller.Database.GetAll(Sql)
        ' Put these in
        For Each Row As DataRow In Table.Rows
            Dim Item As ListViewItem = Me.lvFolders.Items.Add(Row.Item("name"))
            Item.Tag = Row.Item("id")
            Item.ImageIndex = 0
        Next
    End Sub

    Public Sub OpenFolder(ByVal Id As Integer)
        If Id <> Nothing Then
            ' Clear old
            Me.lvDocs.Items.Clear()
            ' Get docs
            Dim Sql As String = "SELECT id, name, template_type_id AS folder_id, date_last_updated, date_created, created_by, last_updated_by"
            Sql &= " FROM template WHERE template_type_id=" & Id
            Me.CurrentFolder = Me.Controller.Database.GetAll(Sql)
            Me.lblTemplates.Text = "Templates in " & Me.lvFolders.SelectedItems(0).Text
            ' Populate docs
            For Each Row As DataRow In Me.CurrentFolder.Rows
                Dim Item As ListViewItem = Me.lvDocs.Items.Add(Row.Item("name"))
                Item.Tag = Row.Item("id")
                Item.ImageIndex = 2
            Next
        End If
    End Sub

    Private Sub SelectAndClose()
        Me.Controller.SelectedFolderID = Me.SelectedFolder
        Me.Controller.SelectedTemplateID = Me.SelectedDocument
        Me.Controller.SelectedTemplateName = Me.txtName.Text
        Me.Controller.SelectedTemplateHtml = Me.SelectedTemplateHtml
        Me.Controller.ButtonPress = Gravity.Response.OK
        Me.Close()
    End Sub

    Private Sub OpenTemplate()
        If Me.lvDocs.SelectedItems.Count > 0 Then
            ' HAS SELECTED AN EXISTING DOCUMENT
            If Me.Controller.WindowMode = Gravity.ChooseTemplateDialog.Mode.SaveTemplate Then
                ' Saving file, ask if overwriting
                Dim Ask As New MyCore.Gravity.AskBox("Overwrite this file?", "Overwrite Confirm")
                If Not Ask.ButtonPress = Gravity.Response.Yes Then
                    Me.SelectAndClose()
                End If
            Else
                Me.SelectAndClose()
            End If
        Else
            ' NOTHING SELECTED
            If Me.Controller.WindowMode = Gravity.ChooseTemplateDialog.Mode.OpenTemplate Then
                Dim ErrorBox As New MyCore.Gravity.ErrorBox("No template selected.")
            Else
                ' Is folder selected?
                If Me.lvFolders.SelectedItems.Count > 0 Then
                    ' Something typed in?
                    If Me.txtName.Text.Trim.Length > 0 Then
                        Me.SelectAndClose()
                    Else
                        Dim ErrorBox As New MyCore.Gravity.ErrorBox("Enter a name of the template.")
                    End If
                Else
                    Dim ErrorBox As New MyCore.Gravity.ErrorBox("Select a folder to put file in first.")
                End If
            End If
        End If
    End Sub

    Private Sub ubtnOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnOK.Click
        Me.OpenTemplate()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnCancel.Click
        Me.Controller.ButtonPress = Gravity.Response.Cancel
        Me.Close()
    End Sub

    Private Sub mnuAddFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddFolder.Click
        Dim Input As New MyCore.Gravity.InputBox("Name of new folder:", "New Folder")
        If Input.ButtonPress = Gravity.Response.OK Then
            Me.Controller.Database.Execute("INSERT INTO form_type (name, sort) VALUES (" & Me.Controller.Database.Escape(Input.Text) & ", 0)")
            If Me.Controller.Database.LastQuery.Successful Then
                Me.LoadFolders()
            Else
                Dim Err As New MyCore.Gravity.ErrorBox(Me.Controller.Database.LastQuery.ErrorMsg)
            End If
        End If
    End Sub

    Private Sub mnuRenameFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRenameFolder.Click
        If Me.SelectedFolder > 0 Then
            Dim Input As New MyCore.Gravity.InputBox
            Input.SetValues("New Name of Folder:", "Rename Folder")
            If Input.ShowDialog() = Gravity.Response.OK Then
                Me.Controller.Database.Execute("UPDATE form_type SET name=" & Me.Controller.Database.Escape(Input.Text) & " WHERE id=" & Me.SelectedFolder)
                If Me.Controller.Database.LastQuery.Successful Then
                    Me.LoadFolders()
                Else
                    Dim Err As New MyCore.Gravity.ErrorBox(Me.Controller.Database.LastQuery.ErrorMsg)
                End If
            End If
        Else
            Dim Info As New MyCore.Gravity.InfoBox("No folder selected.")
        End If
    End Sub

    Private Sub lvDocs_ItemSelectionChanged(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinListView.ItemSelectionChangedEventArgs)
        If e.SelectedItems.Count > 0 Then
            Me.txtName.Text = e.SelectedItems(0).Text
        End If
    End Sub

    Private Sub lvDocs_DoubleClick1(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvDocs.DoubleClick
        Me.OpenTemplate()
    End Sub

    Private Sub mnuRenameDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRenameDoc.Click
        If Me.SelectedDocument > 0 Then
            Dim Input As New MyCore.Gravity.InputBox
            Input.SetValues("New Name of Template:", "Rename Template")
            If Input.ShowDialog() = Gravity.Response.OK Then
                Me.Controller.Database.Execute("UPDATE template SET name=" & Me.Controller.Database.Escape(Input.Text) & " WHERE id=" & Me.SelectedDocument)
                If Me.Controller.Database.LastQuery.Successful Then
                    Me.OpenFolder(Me.SelectedFolder)
                Else
                    Dim Err As New MyCore.Gravity.ErrorBox(Me.Controller.Database.LastQuery.ErrorMsg)
                End If
            End If
        Else
            Dim Info As New MyCore.Gravity.InfoBox("No template selected.")
        End If
    End Sub

    Private Sub mnuDeleteTemplate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDeleteTemplate.Click
        If Me.SelectedDocument > 0 Then
            Dim Ask As New MyCore.Gravity.AskBox("Are you sure you want to delete this template? There is no going back.", "Delete Template", 0)
            If Ask.ButtonPress = Gravity.Response.Yes Then
                Me.Controller.Database.Execute("DELETE FROM template WHERE id=" & Me.SelectedDocument)
                If Me.Controller.Database.LastQuery.Successful Then
                    Me.OpenFolder(Me.SelectedFolder)
                Else
                    Dim Err As New MyCore.Gravity.ErrorBox(Me.Controller.Database.LastQuery.ErrorMsg)
                End If
            End If
        Else
            Dim Info As New MyCore.Gravity.InfoBox("No template selected.")
        End If
    End Sub

    Private Sub DocContext_Opening(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles DocContext.Opening
        If Me.SelectedDocument = Nothing Then
            e.Cancel = True
        End If
    End Sub

    Private Sub lvFolders_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvFolders.Click
        If Me.lvFolders.SelectedItems.Count > 0 Then
            For Each Item As ListViewItem In Me.lvFolders.Items
                Item.ImageIndex = 0
                Item.BackColor = Color.Transparent
                Item.ForeColor = Color.Black
            Next
            Me._SelectedFolder = Me.lvFolders.SelectedItems(0).Tag
            Me.OpenFolder(Me.lvFolders.SelectedItems(0).Tag)
            Me.lvFolders.SelectedItems(0).ImageIndex = 1
        End If
    End Sub

    Private Sub lvFolders_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvFolders.DragDrop
        If e.Data.GetDataPresent("System.Windows.Forms.ListViewItem") Then
            Dim pt As Point = Me.lvFolders.PointToClient(New Point(e.X, e.Y))
            Dim Item As ListViewItem = Me.lvFolders.GetItemAt(pt.X, pt.Y)
            If Item IsNot Nothing Then
                If Item.Tag <> Me.SelectedFolder Then
                    Dim Ask As New MyCore.Gravity.AskBox("Move this " & Me.DragTemplate.Text & " to the " & Item.Text & " folder?", "Move Template", 0)
                    If Ask.ButtonPress = Gravity.Response.Yes Then
                        Me.Controller.Database.Execute("UPDATE template SET template_type_id=" & Item.Tag & " WHERE id=" & Me.DragTemplate.Tag)
                        If Me.Controller.Database.LastQuery.Successful Then
                            Me.OpenFolder(Me.SelectedFolder)
                        Else
                            Dim Err As New MyCore.Gravity.ErrorBox(Me.Controller.Database.LastQuery.ErrorMsg)
                        End If
                    End If
                Else
                    Dim Err As New MyCore.Gravity.ErrorBox("The selected template is already in this folder!")
                End If
            Else
                Dim Info As New MyCore.Gravity.InfoBox("No folder selected.")
            End If
        End If
    End Sub

    Private Sub lvFolders_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvFolders.DragEnter
        If e.Data.GetDataPresent("System.Windows.Forms.ListViewItem") Then
            e.Effect = DragDropEffects.Move
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub lvFolders_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvFolders.DragLeave

    End Sub

    Private Sub lvFolders_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvFolders.DragOver
        Dim pt As Point = Me.lvFolders.PointToClient(New Point(e.X, e.Y))
        Dim Item As ListViewItem = Me.lvFolders.GetItemAt(pt.X, pt.Y)
        If Item IsNot Nothing Then
            For Each li As ListViewItem In Me.lvFolders.Items
                li.BackColor = Color.Transparent
                li.ForeColor = Color.Black
            Next
            Item.Selected = True
            Item.Focused = True
            Item.BackColor = Color.DarkBlue
            Item.ForeColor = Color.White
        End If
    End Sub

    Private Sub lvDocs_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvDocs.DragEnter
        e.Effect = DragDropEffects.None
    End Sub

    Private Sub lvDocs_ItemDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles lvDocs.ItemDrag
        Me.DragTemplate = e.Item
        sender.DoDragDrop(New DataObject("System.Windows.Forms.ListViewItem", e.Item), DragDropEffects.Move)
    End Sub

End Class
