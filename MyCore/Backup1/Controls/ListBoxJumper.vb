Public Class ListBoxJumper
    Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'UserControl overrides dispose to clean up the component list.
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
    Friend WithEvents utxtSearch As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents ugrdList As Infragistics.Win.UltraWinGrid.UltraGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.utxtSearch = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.ugrdList = New Infragistics.Win.UltraWinGrid.UltraGrid
        CType(Me.utxtSearch, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugrdList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'utxtSearch
        '
        Me.utxtSearch.Dock = System.Windows.Forms.DockStyle.Top
        Me.utxtSearch.Location = New System.Drawing.Point(0, 0)
        Me.utxtSearch.Name = "utxtSearch"
        Me.utxtSearch.Size = New System.Drawing.Size(288, 21)
        Me.utxtSearch.TabIndex = 3
        '
        'ugrdList
        '
        Appearance1.BackColorAlpha = Infragistics.Win.Alpha.Opaque
        Me.ugrdList.DisplayLayout.Appearance = Appearance1
        Me.ugrdList.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ResizeAllColumns
        Me.ugrdList.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        Me.ugrdList.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugrdList.DisplayLayout.Override.MaxSelectedRows = 1
        Me.ugrdList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugrdList.Location = New System.Drawing.Point(0, 21)
        Me.ugrdList.Name = "ugrdList"
        Me.ugrdList.Size = New System.Drawing.Size(288, 219)
        Me.ugrdList.TabIndex = 4
        '
        'ListBoxJumper
        '
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.ugrdList)
        Me.Controls.Add(Me.utxtSearch)
        Me.Name = "ListBoxJumper"
        Me.Size = New System.Drawing.Size(288, 240)
        CType(Me.utxtSearch, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugrdList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Event SearchTextChanged(ByVal Text As String)
    Public Event SelectedIndexChanged(ByVal Index As Integer)
    Public Shadows Event DoubleClick(ByVal Value As String)

    Dim _ValueMember As String = ""
    Dim _DisplayMember As String = ""

    Dim blnLoaded As Boolean = False


    Public Property DataSource() As DataTable
        Get
            Return Me.ugrdList.DataSource
        End Get
        Set(ByVal Value As DataTable)
            Me.ugrdList.DataSource = Value
        End Set
    End Property

    Public Property DisplayMember() As String
        Get
            Return Me._DisplayMember
        End Get
        Set(ByVal Value As String)
            Me._DisplayMember = Value
        End Set
    End Property

    Public Property ValueMember() As String
        Get
            Return Me._ValueMember
        End Get
        Set(ByVal Value As String)
            Me._ValueMember = Value
        End Set
    End Property

    Public Property SelectedIndex() As Integer
        Get
            If Me.ugrdList.Selected.Rows.Count > 0 Then
                Return Me.ugrdList.Selected.Rows(0).Index
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal Value As Integer)
            If Me.blnLoaded Then
                Try
                    If Me.ugrdList.Selected.Rows.Count > 0 Then
                        Me.ugrdList.Selected.Rows(0).Selected = False
                    End If
                    Me.ugrdList.Rows(Value).Selected = True
                Catch
                    ' Ignore
                End Try
            End If
        End Set
    End Property

    Public ReadOnly Property DataGrid() As Infragistics.Win.UltraWinGrid.UltraGrid
        Get
            Return Me.ugrdList
        End Get
    End Property

    Private Sub txtSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles utxtSearch.TextChanged
        If Me.blnLoaded Then
            RaiseEvent SearchTextChanged(Me.utxtSearch.Text)
            For i As Integer = 0 To Me.DataSource.Rows.Count - 1
                If Me.ugrdList.Rows(i).Cells(Me._DisplayMember).Text.ToUpper.StartsWith(Me.utxtSearch.Text.ToUpper) Then
                    If Me.ugrdList.Selected.Rows.Count > 0 Then
                        Me.ugrdList.Selected.Rows(0).Selected = False
                    End If
                    Me.ugrdList.Rows(i).Selected = True
                    Me.ugrdList.Rows(i).Activated = True
                    Exit Sub
                Else
                    If Me.ugrdList.Selected.Rows.Count > 0 Then
                        Me.ugrdList.Selected.Rows(0).Selected = False
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub ugrdList_AfterSelectChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles ugrdList.AfterSelectChange
        If Me.blnLoaded Then
            If Me.ugrdList.Selected.Rows.Count = 1 Then
                RaiseEvent SelectedIndexChanged(Me.ugrdList.Selected.Rows(0).Index)
            End If
        End If
    End Sub

    Private Sub ListBoxJumper_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.blnLoaded = True
    End Sub

    Private Sub ugrdList_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles ugrdList.DoubleClickCell
        If Me.blnLoaded Then
            RaiseEvent DoubleClick(Me.ugrdList.Selected.Rows(0).Cells(Me._ValueMember).Value)
        End If
    End Sub

End Class
