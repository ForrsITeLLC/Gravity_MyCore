Public Class AutoCompleteBox

#Region "Properties"

    Dim _TableName As String
    Dim _DisplayMember As String
    Dim _ValueMember As String
    Dim _SearchOn As String
    Dim _WhereClause As String
    Dim _ItemType As String
    Dim _FilterValue As String = ""

    Dim _Win As MyCore.Plugins.Window

    Dim _Hash As New CIHashtable
    Dim _Source As New AutoCompleteStringCollection
    Dim _Value As String = Nothing

    Dim _CanOpen As Boolean = True
    Dim _CanCreateNew As Boolean = True
    Dim _CanSearch As Boolean = True

    Dim _Loaded As Boolean = False
    Dim _Pause As Boolean = False

    Dim _Required As Boolean = False

    Public Property Required() As Boolean
        Get
            Return Me._Required
        End Get
        Set(ByVal value As Boolean)
            If value Then
                Me.TextBox.BackColor = Me._RequiredBackColor
            Else
                Me.TextBox.BackColor = Me._NoMatchBackColor
            End If
            Me._Required = value
        End Set
    End Property

    Dim TextEdited As Boolean = False

    Public Event ValueChanged(ByVal sender As AutoCompleteBox)
    Public Shadows Event TextChanged(ByVal sender As AutoCompleteBox)
    Public Event BeforeFilterAutoComplete(ByVal sender As AutoCompleteBox)
    Public Event BeforeQueryAutoComplete(ByVal sender As AutoCompleteBox, ByVal sql As String)
    Public Event BeforeSearchForMatch(ByVal sender As AutoCompleteBox)
    Public Event OnMatchFound(ByVal sender As AutoCompleteBox)
    Public Event OnNoMatchFound(ByVal sender As AutoCompleteBox)
    Public Event OnLeaveTextBox(ByVal sender As AutoCompleteBox)
    Public Event OnEnterTextBox(ByVal sender As AutoCompleteBox)
    Public Shadows Event OnTextChanged(ByVal sender As AutoCompleteBox)
    Public Shadows Event OnKeyPress(ByVal sender As AutoCompleteBox)

    Public PrePopulate As Boolean = False

    Dim _MinLengthToSearch As Integer = 2

    Public Property MinLengthToSearch() As Integer
        Get
            Return Me._MinLengthToSearch
        End Get
        Set(ByVal value As Integer)
            Me._MinLengthToSearch = value
        End Set
    End Property

    Dim _NoMatchForeColor As Color = Color.DarkRed
    Dim _NoMatchBackColor As Color = Color.FloralWhite
    Dim _NoMatchFont As Font

    Dim _RequiredBackColor As Color = Color.Yellow
    Public Property RequiredBackColor() As Color
        Get
            Return Me._RequiredBackColor
        End Get
        Set(ByVal value As Color)
            Me._RequiredBackColor = value
        End Set
    End Property

    Public Property NoMatchForeColor() As Color
        Get
            Return Me._NoMatchForeColor
        End Get
        Set(ByVal value As Color)
            Me._NoMatchForeColor = value
        End Set
    End Property

    Public Property NoMatchBackColor() As Color
        Get
            Return Me._NoMatchBackColor
        End Get
        Set(ByVal value As Color)
            Me._NoMatchBackColor = value
        End Set
    End Property

    Public Property NoMatchFont() As Font
        Get
            Return Me._NoMatchFont
        End Get
        Set(ByVal value As Font)
            Me._NoMatchFont = value
        End Set
    End Property

    Dim _MatchForeColor As Color = Color.Black
    Dim _MatchBackColor As Color = Color.White
    Dim _MatchFont As Font

    Public Property MatchForeColor() As Color
        Get
            Return Me._MatchForeColor
        End Get
        Set(ByVal value As Color)
            Me._MatchForeColor = value
        End Set
    End Property

    Public Property MatchBackColor() As Color
        Get
             Return Me._MatchBackColor
        End Get
        Set(ByVal value As Color)
            Me._MatchBackColor = value
        End Set
    End Property

    Public Property MatchFont() As Font
        Get
            Return Me._MatchFont
        End Get
        Set(ByVal value As Font)
            Me._MatchFont = value
        End Set
    End Property

    Dim _TypingForeColor As Color = Color.Black
    Dim _TypingBackColor As Color = Color.WhiteSmoke
    Dim _TypingFont As Font

    Public Property TypingForeColor() As Color
        Get
            Return Me._TypingForeColor
        End Get
        Set(ByVal value As Color)
            Me._TypingForeColor = value
        End Set
    End Property

    Public Property TypingBackColor() As Color
        Get
            Return Me._TypingBackColor
        End Get
        Set(ByVal value As Color)
            Me._TypingBackColor = value
        End Set
    End Property

    Public Property TypingFont() As Font
        Get
            Return Me._TypingFont
        End Get
        Set(ByVal value As Font)
            Me._TypingFont = value
        End Set
    End Property

    Dim PresetValue As String = Nothing

    Public Sub Init(ByVal w As MyCore.Plugins.Window)
        Me._Win = w
        Me.SetButtonsNoMatch()
    End Sub

    Public ReadOnly Property Value() As String
        Get
            Return Me._Value
        End Get
    End Property

    Public Overrides Property Text() As String
        Get
            Return Me.TextBox.Text
        End Get
        Set(ByVal Value As String)
            Me.TextBox.Text = Value
            RaiseEvent TextChanged(Me)
        End Set
    End Property

    Public Property ItemType() As String
        Get
            Return Me._ItemType
        End Get
        Set(ByVal value As String)
            Me._ItemType = value
            If value = "Contact" Or value = "Contacts" Then
                Me.PrePopulate = True
            Else
                Me.PrePopulate = False
            End If
        End Set
    End Property

    Public Property FilterValue() As String
        Get
            Return Me._FilterValue
        End Get
        Set(ByVal value As String)
            Me._FilterValue = value
            If Me.PrePopulate Then
                Me.GetSource()
            End If
        End Set
    End Property

    Public ReadOnly Property ButtonWidth() As Integer
        Get
            Dim Width As Integer = 0
            If Me.btnNew.Visible Then
                Width += Me.btnNew.Width
            End If
            If Me.btnOpen.Visible Then
                Width += Me.btnOpen.Width
            End If
            If Me.btnSearch.Visible Then
                Width += Me.btnSearch.Width
            End If
            Return Width
        End Get
    End Property

    Private Function IsPaused() As Boolean
        Return Me._Pause
    End Function

    Private Sub Pause()
        Me._Pause = True
    End Sub

    Private Sub Unpause()
        Me._Pause = False
    End Sub

#End Region

    Private Sub LookupBox_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Me.TypingFont Is Nothing Then
            Me.TypingFont = New Font(Me.TextBox.Font, FontStyle.Regular)
        End If
        If Me.NoMatchFont Is Nothing Then
            Me.NoMatchFont = New Font(Me.TextBox.Font, FontStyle.Italic)
        End If
        If Me.MatchFont Is Nothing Then
            Me.MatchFont = New Font(Me.TextBox.Font, FontStyle.Bold)
        End If
        Select Case Me._ItemType
            Case "Customer", "Customers"
                Me._TableName = "ADDRESS"
                Me._DisplayMember = "cst_name"
                Me._ValueMember = "cst_no"
                Me._SearchOn = "cst_name"
                Me._WhereClause = "type IN (SELECT id FROM company_type WHERE vendor=0)"
            Case "Vendor", "Vendors"
                Me._TableName = "ADDRESS"
                Me._DisplayMember = "cst_name"
                Me._ValueMember = "cst_no"
                Me._SearchOn = "cst_name"
                Me._WhereClause = "type IN (SELECT id FROM company_type WHERE vendor=1)"
            Case "Company", "Companies"
                Me._TableName = "ADDRESS"
                Me._DisplayMember = "cst_name"
                Me._ValueMember = "cst_no"
                Me._SearchOn = "cst_name"
                Me._WhereClause = Nothing
            Case "Contact", "Contacts"
                Me._TableName = "CONTACTS"
                Me._DisplayMember = "ISNULL(cnt_last, '') + ', ' + ISNULL(cnt_first, '')"
                Me._ValueMember = "cnt_id"
                Me._SearchOn = "cnt_last + ', ' + cnt_first"
                Me._WhereClause = "cnt_no='%filter%'"
        End Select
        Me._Loaded = True
        Me.TextBox.BackColor = Me._NoMatchBackColor
        If Me.PresetValue <> Nothing Then
            Me.SetValue(Me.PresetValue)
        Else
            Me.SetNoMatch()
        End If
    End Sub

    Private Sub AutoCompleteBox_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

    End Sub

    Private Sub GetSource()
        RaiseEvent BeforeFilterAutoComplete(Me)
        Me.ClearSouce()
        Dim Sql As String = "SELECT " & Me._DisplayMember & " AS displayname, " & Me._ValueMember & " AS id"
        ' Hack for contacts
        If Me._ItemType = "Contact" Or Me._ItemType = "Contacts" Then
            Sql &= ", ISNULL(cnt_first, '') + ' ' + ISNULL(cnt_last, '') AS displayname2"
        End If
        Sql &= " FROM " & Me._TableName
        If Me._WhereClause <> Nothing Or Me.Text.Length > 0 Then
            Sql &= " WHERE "
            If Me._WhereClause <> Nothing Then
                Sql &= Me._WhereClause.Replace("%filter%", Me.FilterValue)
            End If
            If Me.Text.Length > 0 Then
                If Me._WhereClause <> Nothing Then
                    Sql &= " AND"
                End If
                Sql &= " " & Me._SearchOn & " LIKE " & Me._Win.ParentWin.Database.Escape(Me.Text & "%")
            End If
        End If
        Sql &= " ORDER BY " & Me._DisplayMember
        RaiseEvent BeforeQueryAutoComplete(Me, Sql)
        Dim Table As DataTable = Me._Win.ParentWin.Database.GetAll(Sql)
        If Me._Win.ParentWin.Database.LastQuery.Successful Then
            If Table.Rows.Count > 0 Then
                Me._Source = New AutoCompleteStringCollection
                For Each Row As DataRow In Table.Rows
                    Try
                        If Not Me._Hash.ContainsKey(Row.Item("displayname").ToString.Trim) Then
                            Me.AddToDataSource(Row.Item("displayname").ToString.Trim, Row.Item("id"))
                        End If
                        ' Hack for contacts
                        If Me._ItemType = "Contact" Or Me._ItemType = "Contacts" Then
                            If Not Me._Hash.ContainsKey(Row.Item("displayname2").ToString.Trim) Then
                                Me.AddToDataSource(Row.Item("displayname2").ToString.Trim, Row.Item("id"))
                            End If
                        End If
                    Catch
                        ' error can happen when displayname is blank.  Just ignore it and skip this person.
                    End Try
                Next
                Me.TextBox.AutoCompleteCustomSource = Me._Source
            Else
                Me.TextBox.AutoCompleteCustomSource = New AutoCompleteStringCollection
            End If
        Else
            Me._Win.ParentWin.Err.ShowDialog(Me._Win.ParentWin.Database.LastQuery.ErrorMsg, "Error", Sql)
        End If
    End Sub

    Public ReadOnly Property Resolved() As Boolean
        Get
            If Me._Value <> Nothing Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public Sub AddToDataSource(ByVal Display As String, ByVal Value As String)
        If Not Me._Hash.ContainsKey(Display) Then
            Me._Source.Add(Display)
            Me._Hash.Add(Display, Value)
        End If
    End Sub

    Public Sub ClearSouce()
        Dim Source As New AutoCompleteStringCollection
        Me.TextBox.AutoCompleteSource = AutoCompleteSource.CustomSource
        Me.TextBox.AutoCompleteCustomSource = Source
        Me._Hash = New CIHashtable
    End Sub

    Public Function GetValue(ByVal IfEmptyValue As Object) As Object
        If Me.Resolved Then
            Return Me.Value
        Else
            Return IfEmptyValue
        End If
    End Function

    Public Sub SetValue(ByVal Value As String)
        Me.TextBox.Text = Value
        Me.IsMatch(True)
    End Sub

    Public Sub SetMatch(ByVal Display As String, ByVal Value As String)
        Me.Pause()
        Me.TextBox.Text = Display
        Me._Value = Value
        Me.ToolTip1.SetToolTip(Me.TextBox, Display & " (" & Value & ")")
        Me.SetFontMatched()
        Me.SetButtonsMatch()
        Me.Unpause()
        RaiseEvent ValueChanged(Me)
        RaiseEvent OnMatchFound(Me)
    End Sub

    Private Sub SetNoMatch(Optional ByVal Display As String = "")
        Me._Value = Nothing
        If Display.Length > 0 Then
            Me.ToolTip1.SetToolTip(Me.TextBox, "No match for " & Display)
        Else
            Me.ToolTip1.SetToolTip(Me.TextBox, "Nothing")
        End If
        Me.SetFontNoMatch()
        Me.SetButtonsNoMatch()
        RaiseEvent ValueChanged(Me)
        RaiseEvent OnNoMatchFound(Me)
    End Sub

    Public Function IsMatch(Optional ByVal IsValue As Boolean = False) As Boolean
        RaiseEvent BeforeSearchForMatch(Me)
        If Me.TextBox.Text.Trim.Length > 0 Then
            Dim Input As String = Me.TextBox.Text.Trim
            If Me._Hash.ContainsKey(Input) And Not IsValue Then
                ' Found match in source
                Me.SetMatch(Input, Me._Hash.Item(Input))
                Return True
            Else
                ' Maybe it is an id #, look it up
                Dim Sql As String = "SELECT " & Me._DisplayMember & " AS displayname FROM " & Me._TableName
                Sql &= " WHERE " & Me._ValueMember & "=" & Me._Win.ParentWin.Database.Escape(Input)
                'If Me._WhereClause <> Nothing Then
                '    Sql &= " AND " & Me._WhereClause
                'End If
                Dim Display As String = Me._Win.ParentWin.Database.GetOne(Sql)
                If Display <> Nothing Then
                    Me.SetMatch(Display, Input)
                    Return True
                Else
                    ' Maybe it is a name that wasn't in hash for some reason, look it up
                    Sql = "SELECT " & Me._ValueMember & " AS value FROM " & Me._TableName
                    Sql &= " WHERE " & Me._DisplayMember & "=" & Me._Win.ParentWin.Database.Escape(Input)
                    If Me._WhereClause <> Nothing Then
                        Sql &= " AND " & Me._WhereClause
                    End If
                    Dim Value As String = Me._Win.ParentWin.Database.GetOne(Sql)
                    If Value <> Nothing Then
                        Me.SetMatch(Input, Value)
                        Return True
                    Else
                        Me.SetNoMatch(Input)
                        Return False
                    End If
                End If
            End If
        Else
            Me.SetNoMatch()
            Return False
        End If
    End Function

    Private Sub SetFontNoMatch()
        Me.TextBox.ForeColor = Me.NoMatchForeColor
        Me.TextBox.BackColor = Me.GetBackColor
        Me.TextBox.Font = Me.NoMatchFont
    End Sub

    Private Sub SetFontMatched()
        Me.TextBox.ForeColor = Me.MatchForeColor
        Me.TextBox.BackColor = Me.GetBackColor
        Me.TextBox.Font = Me.MatchFont
    End Sub

    Private Sub SetFontTyping()
        Me.TextBox.ForeColor = Me.TypingForeColor
        Me.TextBox.BackColor = Me.TypingBackColor
        Me.TextBox.Font = Me.TypingFont
    End Sub

    Private Function GetBackColor() As Drawing.Color
        If Me.Required Then
            Return Me._RequiredBackColor
        ElseIf Me.Resolved Then
            Return Me._MatchBackColor
        Else
            Return Me._NoMatchBackColor
        End If
    End Function

    Private Sub AdjustSize()
        Me.SuspendLayout()
        Me.pnlTextBox.Width = Me.Width - Me.ButtonWidth
        Me.ResumeLayout()
        Me.Refresh()
    End Sub

    Private Sub SetButtonsNoMatch()
        Me.SuspendLayout()
        Me.btnOpen.Visible = False
        Me.btnNew.Visible = True
        Me.ResumeLayout()
        Me.Refresh()
        Me.AdjustSize()
    End Sub

    Private Sub SetButtonsHidden()
        Me.SuspendLayout()
        Me.btnOpen.Visible = False
        Me.btnNew.Visible = False
        Me.ResumeLayout()
        Me.Refresh()
        Me.AdjustSize()
    End Sub

    Private Sub SetButtonsMatch()
        Me.SuspendLayout()
        Me.btnOpen.Visible = True
        Me.btnNew.Visible = False
        Me.ResumeLayout()
        Me.Refresh()
        Me.AdjustSize()
    End Sub

#Region "Clicks"

    Private Sub ubtnOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpen.Click
        If Me.Resolved Then
            Dim Win As MyCore.Plugins.Window = Me._Win.ParentWin.CallMethod(Me, Plugins.MethodType.Open, Me._ItemType, Me.Value)
            Win.Open()
        End If
    End Sub

    Private Sub ubtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        If Me.ItemType = "Contact" Then
            Dim Value As String = Me.FilterValue
            If Value Is DBNull.Value Or Value = Nothing Then
                Value = ""
            End If
            If Value.Length > 0 Then
                Dim Win As MyCore.Plugins.Window = Me._Win.ParentWin.CallMethod(Me, Plugins.MethodType.Lookup, Me._ItemType, Value)
                AddHandler Win.OnEvent, AddressOf Me.SearchMatch
                Win.Open()
            End If
        Else
            Dim Value As String = Me.Value
            If Value Is DBNull.Value Or Value = Nothing Then
                Value = ""
            End If
            Dim Win As MyCore.Plugins.Window = Me._Win.ParentWin.CallMethod(Me, Plugins.MethodType.Lookup, Me._ItemType, Value)
            AddHandler Win.OnEvent, AddressOf Me.SearchMatch
            Win.Open()
        End If
    End Sub

    Private Sub SearchMatch(ByVal Win As MyCore.Plugins.Window, ByVal e As MyCore.Plugins.Window.EventInfo)
        If e.WhatHappened = Plugins.EventType.ItemOpened Then
            Me.SetMatch(e.Item2, e.Item1)
        End If
    End Sub

    Private Sub ubtnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim Win As MyCore.Plugins.Window
        Win = Me._Win.ParentWin.CallMethod(Me, Plugins.MethodType.NewItem, Me._ItemType, Me.Text, "Name", Me.FilterValue, "Filter")
        If Win IsNot Nothing Then
            AddHandler Win.OnEvent, AddressOf Me.NewCreated
            Win.Open()
        End If
    End Sub

    Private Sub NewCreated(ByVal Win As MyCore.Plugins.Window, ByVal e As MyCore.Plugins.Window.EventInfo)
        If e.WhatHappened = Plugins.EventType.NewItemCreated Then
            If e.Item1Arg = "Company" Then
                Dim Item As MyCore.cCompany = e.Item1
                Me.SetMatch(Item.Name, Item.CustomerNo)
            ElseIf e.Item1Arg = "Part No" Then
                Dim Item As MyCore.cItemMaster = e.Item1
                Me.SetMatch(Item.Name, Item.PartNo)
            End If
        End If
    End Sub

#End Region

    Private Sub TextBox_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox.KeyUp
        If Not Me.IsPaused And Not Me.PrePopulate Then
            Me.Pause()
            If e.KeyCode = Keys.Tab Or e.KeyCode = Keys.Enter Then
                'MessageBox.Show(e.KeyCode)
            Else
                RaiseEvent OnKeyPress(Me)
                If Me.TextBox.Text.Length >= Me.MinLengthToSearch Then
                    If Me._Hash.Count = 0 Then
                        Me.GetSource()
                    End If
                Else
                    Me.ClearSouce()
                End If
            End If
            Me.Unpause()
        End If
    End Sub

    Private Sub TextBox_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox.LostFocus
        If Not Me.IsPaused And Me.TextEdited Then
            Me.Pause()
            Me.IsMatch()
            Me.Unpause()
            RaiseEvent OnLeaveTextBox(Me)
        End If
    End Sub

    Private Sub TextBox_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox.Leave
        Me.TextBox.BackColor = Me.GetBackColor
    End Sub

    Private Sub TextBox_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox.Enter
        If Not Me.IsPaused Then
            RaiseEvent OnEnterTextBox(Me)
            Me.Pause()
            Me.SetFontTyping()
            Me.SetButtonsHidden()
            Me.ToolTip1.RemoveAll()
            Me.Unpause()
        End If
    End Sub

    Private Sub TextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox.TextChanged
        If Not Me.IsPaused Then
            TextEdited = True
            RaiseEvent OnTextChanged(Me)
        End If
    End Sub

    Private Sub btnNew_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Resize
        Me.Refresh()
    End Sub

    Private Sub AutoCompleteBox_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize, Me.Validated, Me.Move, Me.DockChanged
        Me.AdjustSize()
    End Sub

End Class
