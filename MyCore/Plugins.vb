Imports System.IO
Imports System.Reflection

Namespace Plugins

    Public Enum EventType
        NewItemCreated = 1
        ItemOpened = 2
        GenericMessage = 3
        ChangesSaved = 4
        ErrorCreatingItem = -1
        ErrorOpeningItem = -2
        GenericErrorMessage = -3
        ErrorSavingChanges = -4
    End Enum

    Public Enum MethodType
        Open = 1
        NewItem = 2
        Search = 3
        Lookup = 4
        Browse = 5
        Merge = 6
        Rename = 7
        Report = 8
        Delete = 9
        Execute = 10
    End Enum

    Public Class Plugins

        Public Structure AvailablePlugin
            Public AssemblyPath As String
            Public ClassName As String
        End Structure

        Public Shared Function FindPlugins(ByVal path As String, ByVal strInterface As String) As AvailablePlugin()

            Dim Plugins As New ArrayList()
            Dim Dll As [Assembly]

            For Each file As String In My.Computer.FileSystem.GetFiles(path, FileIO.SearchOption.SearchTopLevelOnly, "*.dll")
                Try
                    Dll = [Assembly].LoadFrom(file)
                    ExamineAssembly(Dll, strInterface, Plugins)
                Catch ex As Exception

                End Try
            Next

            Dim Results(Plugins.Count - 1) As AvailablePlugin

            If Plugins.Count <> 0 Then
                Plugins.CopyTo(Results)
                Return Results
            Else
                Return Nothing
            End If

        End Function

        Private Shared Sub ExamineAssembly(ByVal dll As [Assembly], ByVal strInterface As String, ByVal plugins As ArrayList)
            Dim objInterace As Type
            Dim Plugin As AvailablePlugin
            For Each objType As Type In dll.GetTypes
                If objType.IsPublic Then
                    If Not ((objType.Attributes And TypeAttributes.Abstract) = TypeAttributes.Abstract) Then
                        Dim Name As String = objType.Name
                        objInterace = objType.GetInterface(strInterface, True)
                        If Not (objInterace Is Nothing) Then
                            Plugin = New AvailablePlugin()
                            Plugin.AssemblyPath = dll.Location
                            Plugin.ClassName = objType.FullName
                            plugins.Add(Plugin)
                        End If
                    End If
                End If
            Next
        End Sub

        Public Shared Function CreateInstance(ByVal Plugin As AvailablePlugin) As Object
            Dim dll As [Assembly]
            Try
                dll = [Assembly].LoadFrom(Plugin.AssemblyPath)
                Return dll.CreateInstance(Plugin.ClassName)
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

    End Class

    Public Interface IReport

        Property MapLongitudeField() As String
        Property MapLatitudeField() As String
        Property MapLabelField() As String
        Property MapDescriptionField() As String
        Property Title() As String
        Property Table() As DataTable
        Property CanOpen() As Boolean
        Property ContentItemType() As String
        Property PrimaryKey() As String
        Property DefaultFilters() As Hashtable

        Sub GetReport(ByVal Table As DataTable)
        Sub GetReport(ByVal Sql As String)
        Sub GetReport(ByVal Table As DataView)
        Sub ColorAllLessThan(ByVal BgColor As Drawing.Color, ByVal FieldName As String, ByVal Value As Object)
        Sub ColorAllGreaterThan(ByVal BgColor As Drawing.Color, ByVal FieldName As String, ByVal Value As Object)
        Sub AddHiddenColumn(ByVal Key As String, ByVal Header As String, ByVal Width As Integer, Optional ByVal RowFiltering As Boolean = True, Optional ByVal Format As String = "", Optional ByVal Style As Infragistics.Win.UltraWinGrid.ColumnStyle = Infragistics.Win.UltraWinGrid.ColumnStyle.Default)
        Sub AddColumn(ByVal Key As String, ByVal Header As String, ByVal Width As Integer, Optional ByVal RowFiltering As Boolean = True, Optional ByVal Format As String = "", Optional ByVal Style As Infragistics.Win.UltraWinGrid.ColumnStyle = Infragistics.Win.UltraWinGrid.ColumnStyle.Default)
        Sub AddToolTip(ByVal HoverField As String, ByVal MessageField As String, Optional ByVal Title As String = "Note")
        Sub HideAllColumns()
        Sub ShowAllColumns()
        Sub AddSummary(ByVal Key As String, ByVal Type As Infragistics.Win.UltraWinGrid.SummaryType, ByVal SourceCol As String, ByVal SummaryPosition As Infragistics.Win.UltraWinGrid.SummaryPosition, Optional ByVal Format As String = "")


    End Interface

    Public Interface IMap

        Event MapLoaded(ByVal Win As IMap)

        Property Markers() As Marker()

        Sub CenterAndZoom(ByVal lon As Double, ByVal lat As Double, ByVal zoom As Integer)
        Sub CenterAndZoom()
        Sub DrawMarkers()

        Class Marker

            Public Lon As String
            Public Lat As String
            Public Label As String
            Public Description As String

            Public Sub New(ByVal Lat As String, ByVal Lon As String, Optional ByVal Label As String = "", Optional ByVal Desc As String = "")
                Me.Lat = Lat
                Me.Lon = Lon
                Me.Label = Label
                Me.Description = Desc
            End Sub

        End Class

    End Interface

    Public Interface IPlugin

        Sub Initialize(ByRef ParentWindow As IHost)

        Enum MenuLocation
            Toolbar1 = 0
            Toolbar2 = 1
            Menubar = 2
        End Enum

        Function CallMethod(ByVal Method As MethodType, ByVal ItemType As String, Optional ByRef Value1 As Object = Nothing, Optional ByVal Arg1 As String = Nothing, Optional ByRef Value2 As Object = Nothing, Optional ByVal Arg2 As String = Nothing) As Window

        ReadOnly Property Name() As String
        ReadOnly Property MenuItem() As Object
        ReadOnly Property MenuItemLocation() As MenuLocation
        ReadOnly Property MenuItemIndex() As Integer
        ReadOnly Property PermissionLevel() As Integer

    End Interface

    Public Interface IHost

        Property Database() As MyCore.Data.EasySql
        Property CurrentUser() As MyCore.cEmployee
        ReadOnly Property SettingsGlobal() As MyCore.cSettings
        Property SettingsLocal() As MyCore.Settings.AppSettings
        ReadOnly Property Config() As MyCore.Settings.ConfigFile
        ReadOnly Property Err() As MyCore.Gravity.ErrorBox
        ReadOnly Property Input() As MyCore.Gravity.InputBox
        ReadOnly Property Info() As MyCore.Gravity.InfoBox
        ReadOnly Property Ask() As MyCore.Gravity.AskBox

        ' Call a method, return the window item or nothing
        Function CallMethod(ByRef Sender As Object, ByVal Method As MethodType, ByVal ItemType As String, Optional ByVal Value1 As Object = Nothing, Optional ByVal Arg1 As String = Nothing, Optional ByVal Value2 As Object = Nothing, Optional ByVal Arg2 As String = Nothing) As Window
        Function CallMethod(ByRef Sender As Object, ByVal ItemType As String, ByVal Value As Object) As Window

        ' Open IWindow
        Sub OpenWindow(ByRef Win As Window, Optional ByVal Dialog As Boolean = False)

        ' Plugin Calls this to register itself as a handler
        Sub RegisterMethod(ByVal MethodName As MethodType, ByVal Type As String, ByRef Plugin As IPlugin)

        ' History
        Sub AddToHistory(ByVal ItemType As String, ByVal Id As String, ByVal Display As String)

    End Interface

    Public Class Host
        Inherits Windows.Forms.Form
        Implements MyCore.Plugins.IHost

        Public _db As MyCore.Data.EasySql
        Public _SettingsGlobal As MyCore.cSettings
        Public _SettingsLocal As New MyCore.Settings.AppSettings
        Public _Config As MyCore.Settings.ConfigFile
        Public _Err As New MyCore.Gravity.ErrorBox
        Public _Info As New MyCore.Gravity.InfoBox
        Public _Input As New MyCore.Gravity.InputBox
        Public _Ask As New MyCore.Gravity.AskBox
        Public _MyUser As MyCore.cEmployee
        Public _DemoMode As Boolean = False
        Public _FileName As String = ""
        Public _ProgramName As String = ""

        Public IsLoaded As Boolean = False
        Public Plugins As New Hashtable
        Public Functions As New Hashtable
        Public WithEvents History As New HistoryList

        Public Event Do_SetDatabaseConString(ByVal Str As String)
        Public Event Do_OpenWindow(ByVal Win As MyCore.Plugins.Window, ByVal Dialog As Boolean)
        Public Event Do_CallMethod(ByRef Sender As Object, ByVal Method As MyCore.Plugins.MethodType, ByVal ItemType As String, ByVal Value1 As Object, ByVal Arg1 As String, ByVal Value2 As Object, ByVal Arg2 As String)
        Public Event Do_RegisterMethod(ByVal Method As MyCore.Plugins.MethodType, ByVal Type As String, ByRef Plugin As IPlugin)
        Public Event Do_PluginProgressUpdate(ByVal Plugin As Object, ByVal Current As Integer, ByVal Total As Integer, ByVal Message As String)

        Public ReadOnly Property Config() As MyCore.Settings.ConfigFile Implements MyCore.Plugins.IHost.Config
            Get
                Return Me._Config
            End Get
        End Property

        Public Property Database() As MyCore.Data.EasySql Implements MyCore.Plugins.IHost.Database
            Get
                Return Me._db
            End Get
            Set(ByVal value As MyCore.Data.EasySql)
                Me._db = value
            End Set
        End Property

        Public ReadOnly Property SettingsGlobal() As MyCore.cSettings Implements MyCore.Plugins.IHost.SettingsGlobal
            Get
                Return Me._SettingsGlobal
            End Get
        End Property

        Public Property SettingsLocal() As MyCore.Settings.AppSettings Implements MyCore.Plugins.IHost.SettingsLocal
            Get
                Return Me._SettingsLocal
            End Get
            Set(ByVal value As MyCore.Settings.AppSettings)
                Me._SettingsLocal = value
            End Set
        End Property

        Public ReadOnly Property Ask() As MyCore.Gravity.AskBox Implements MyCore.Plugins.IHost.Ask
            Get
                Return Me._Ask
            End Get
        End Property

        Public ReadOnly Property Err() As MyCore.Gravity.ErrorBox Implements MyCore.Plugins.IHost.Err
            Get
                Return Me._Err
            End Get
        End Property

        Public ReadOnly Property Info() As MyCore.Gravity.InfoBox Implements MyCore.Plugins.IHost.Info
            Get
                Return Me._Info
            End Get
        End Property

        Public ReadOnly Property Input() As MyCore.Gravity.InputBox Implements MyCore.Plugins.IHost.Input
            Get
                Return Me._Input
            End Get
        End Property

        Public Property CurrentUser() As MyCore.cEmployee Implements IHost.CurrentUser
            Get
                Return Me._MyUser
            End Get
            Set(ByVal value As MyCore.cEmployee)
                Me._MyUser = value
            End Set
        End Property

        Public Property ProgramName() As String
            Get
                Return Me._ProgramName
            End Get
            Set(ByVal value As String)
                Me._ProgramName = value
            End Set
        End Property

        Delegate Sub OpenWindowCallback(ByRef Win As MyCore.Plugins.Window, ByVal Dialog As Boolean)

        Public Overridable Sub OpenWindow(ByRef Win As MyCore.Plugins.Window, Optional ByVal Dialog As Boolean = False) Implements MyCore.Plugins.IHost.OpenWindow
            If InvokeRequired Then
                Invoke(New OpenWindowCallback(AddressOf Me.OpenWindow), New Object() {Win, Dialog})
            Else
                If Win IsNot Nothing Then
                    Try
                        Win.ParentWin = Me
                        Win.Owner = Me
                        If Dialog Then
                            RaiseEvent Do_OpenWindow(Win, Dialog)
                            Win.ShowDialog()
                        Else
                            Win.Form.MdiParent = Me
                            Dim Thread As New System.Threading.Thread(AddressOf Win.Show)
                            Thread.Start()
                            RaiseEvent Do_OpenWindow(Win, Dialog)
                        End If
                    Catch ex As Exception
                        Me.Err.ShowDialog("Error opening window.", "Error", ex.ToString)
                    End Try
                End If
            End If
        End Sub

        Delegate Sub RegisterMethodCallback(ByVal Method As MyCore.Plugins.MethodType, ByVal Type As String, ByRef Plugin As IPlugin)

        Public Overridable Sub RegisterMethod(ByVal Method As MyCore.Plugins.MethodType, ByVal Type As String, ByRef Plugin As IPlugin) Implements MyCore.Plugins.IHost.RegisterMethod
            If InvokeRequired Then
                Invoke(New RegisterMethodCallback(AddressOf Me.RegisterMethod), New Object() {Method, Type, Plugin})
            Else
                Dim strMethod As String = "Open"
                If Method = MethodType.NewItem Then
                    strMethod = "New"
                ElseIf Method = MethodType.Search Then
                    strMethod = "Search"
                ElseIf Method = MethodType.Lookup Then
                    strMethod = "Lookup"
                ElseIf Method = MethodType.Browse Then
                    strMethod = "Browse"
                ElseIf Method = MethodType.Delete Then
                    strMethod = "Delete"
                ElseIf Method = MethodType.Execute Then
                    strMethod = "Execute"
                ElseIf Method = MethodType.Merge Then
                    strMethod = "Merge"
                ElseIf Method = MethodType.Rename Then
                    strMethod = "Rename"
                ElseIf Method = MethodType.Report Then
                    strMethod = "Report"
                End If
                If Me.Functions.ContainsKey(strMethod & " " & Type) Then
                    Me.Functions(strMethod & " " & Type) = Plugin
                Else
                    Me.Functions.Add(strMethod & " " & Type, Plugin)
                End If
                RaiseEvent Do_RegisterMethod(Method, Type, Plugin)
            End If
        End Sub

        Public Overridable Sub SetDatabaseConString(ByVal str As String)
            'Dim Type As String = str.Substring(0, str.IndexOf(":"))
            'str = str.Substring(str.IndexOf("://") + 3)
            Me._db = New MyCore.Data.EasySql(str, "mssql")
            Me._SettingsGlobal = New MyCore.cSettings(Me._db)
            RaiseEvent Do_SetDatabaseConString(str)
        End Sub

        Public Overridable Function GetHandler(ByVal Method As MyCore.Plugins.MethodType, ByVal ItemType As String) As MyCore.Plugins.IPlugin
            Dim strMethod As String = "Open"
            If Method = MethodType.NewItem Then
                strMethod = "New"
            ElseIf Method = MethodType.Search Then
                strMethod = "Search"
            ElseIf Method = MethodType.Lookup Then
                strMethod = "Lookup"
            ElseIf Method = MethodType.Browse Then
                strMethod = "Browse"
            ElseIf Method = MethodType.Report Then
                strMethod = "Report"
            ElseIf Method = MethodType.Delete Then
                strMethod = "Delete"
            ElseIf Method = MethodType.Merge Then
                strMethod = "Merge"
            ElseIf Method = MethodType.Rename Then
                strMethod = "Rename"
            ElseIf Method = MethodType.Execute Then
                strMethod = "Execute"
            End If
            If Me.Functions.ContainsKey(strMethod & " " & ItemType) Then
                Return Me.Functions(strMethod & " " & ItemType)
            Else
                Return Nothing
            End If
        End Function

        Delegate Function CallMethodCallback(ByRef Sender As Object, ByVal Method As MyCore.Plugins.MethodType, ByVal ItemType As String, ByVal Value1 As Object, ByVal Arg1 As String, ByVal Value2 As Object, ByVal Arg2 As String) As MyCore.Plugins.Window

        Public Overridable Function CallMethod(ByRef Sender As Object, ByVal Method As MyCore.Plugins.MethodType, ByVal ItemType As String, Optional ByVal Value1 As Object = Nothing, Optional ByVal Arg1 As String = Nothing, Optional ByVal Value2 As Object = Nothing, Optional ByVal Arg2 As String = Nothing) As MyCore.Plugins.Window Implements MyCore.Plugins.IHost.CallMethod
            If InvokeRequired Then
                Return Invoke(New CallMethodCallback(AddressOf Me.CallMethod), New Object() {Sender, Method, ItemType, Value1, Arg1, Value2, Arg2})
            Else
                Dim Plugin As IPlugin = Me.GetHandler(Method, ItemType)
                If Plugin IsNot Nothing Then
                    If Plugin.PermissionLevel <= Me._MyUser.Permission Then
                        Try
                            Dim Win As Window = Plugin.CallMethod(Method, ItemType, Value1, Arg1, Value2, Arg2)
                            If Win IsNot Nothing Then
                                Win.Opener = Sender
                                Win.ParentWin = Me
                            End If
                            RaiseEvent Do_CallMethod(Sender, Method, ItemType, Value1, Arg1, Value2, Arg2)
                            Return Win
                        Catch ex As Exception
                            Me.Err.ShowDialog("There was an error trying to open the window.", "Error", ex.ToString)
                            Return Nothing
                        End Try
                    Else
                        Me.Err.ShowDialog("You do not have permission to access the " & Plugin.Name & " plugin.", "Permission Denied")
                        Return Nothing
                    End If
                Else
                    Me.Err.ShowDialog("There was no plugin registered to handle " & Method.ToString & " " & ItemType)
                    Return Nothing
                End If
            End If
        End Function

        Delegate Function CallMethodCallback2(ByRef Sender As Object, ByVal ItemType As String, ByVal Value As Object) As MyCore.Plugins.Window

        Public Overridable Function CallMethod(ByRef Sender As Object, ByVal ItemType As String, ByVal Value As Object) As MyCore.Plugins.Window Implements MyCore.Plugins.IHost.CallMethod
            If InvokeRequired Then
                Return Invoke(New CallMethodCallback2(AddressOf Me.CallMethod), New Object() {Sender, ItemType, Value})
            Else
                Return Me.CallMethod(Sender, MethodType.Open, ItemType, Value)
            End If
        End Function

        Delegate Function CallWindowlessMethodCallback(ByRef Sender As Object, ByVal Method As MyCore.Plugins.MethodType, ByVal ItemType As String, ByVal Value1 As Object, ByVal Arg1 As String, ByVal Value2 As Object, ByVal Arg2 As String) As Object

        Public Overridable Function CallWindowlessMethod(ByRef Sender As Object, ByVal Method As MyCore.Plugins.MethodType, ByVal ItemType As String, Optional ByVal Value1 As Object = Nothing, Optional ByVal Arg1 As String = Nothing, Optional ByVal Value2 As Object = Nothing, Optional ByVal Arg2 As String = Nothing) As Object
            If InvokeRequired Then
                Return Invoke(New CallWindowlessMethodCallback(AddressOf Me.CallMethod), New Object() {Sender, Method, ItemType, Value1, Arg1, Value2, Arg2})
            Else
                Dim Plugin As IPlugin = Me.GetHandler(Method, ItemType)
                If Plugin IsNot Nothing Then
                    If Plugin.PermissionLevel <= Me._MyUser.Permission Then
                        Try
                            Plugin.CallMethod(Method, ItemType, Value1, Arg1, Value2, Arg2)
                            RaiseEvent Do_CallMethod(Sender, Method, ItemType, Value1, Arg1, Value2, Arg2)
                            Return Plugin
                        Catch ex As Exception
                            Me.Err.ShowDialog("There was an error trying to run windowless method.", "Error", ex.ToString)
                            Return Nothing
                        End Try
                    Else
                        Me.Err.ShowDialog("You do not have permission to access the " & Plugin.Name & " plugin.", "Permission Denied")
                        Return Nothing
                    End If
                Else
                    Me.Err.ShowDialog("There was no plugin registered to handle " & Method.ToString & " " & ItemType)
                    Return Nothing
                End If
            End If
        End Function

        Public Overridable Sub AddToHistory(ByVal ItemType As String, ByVal Id As String, ByVal Display As String) Implements IHost.AddToHistory
            Me.History.Add(ItemType, Id, Display)
        End Sub

        Private Sub RestoreWindowState()
            Dim LastState As System.Windows.Forms.FormWindowState = Me.SettingsLocal.GetSetting(Me.ProgramName & "Window Last State", FormWindowState.Maximized)
            If LastState = FormWindowState.Normal Then
                Me.WindowState = FormWindowState.Normal
                Me.Left = Me.SettingsLocal.GetSetting(Me.ProgramName & "Window Last Left", 0)
                Me.Top = Me.SettingsLocal.GetSetting(Me.ProgramName & "Window Last Top", 0)
                Me.Width = Me.SettingsLocal.GetSetting(Me.ProgramName & "Window Last Width", 1000)
                Me.Height = Me.SettingsLocal.GetSetting(Me.ProgramName & "Window Last Height", 600)
            Else
                Me.WindowState = FormWindowState.Maximized
            End If
        End Sub

        Private Sub Host_Loaded(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
            Me.RestoreWindowState()
            Me.IsLoaded = True
        End Sub

        Private Sub Host_Closing(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.FormClosing
            Me.SettingsLocal.SaveSetting(Me.ProgramName & "Window Last Left", Me.Left)
            Me.SettingsLocal.SaveSetting(Me.ProgramName & "Window Last Top", Me.Top)
            Me.SettingsLocal.SaveSetting(Me.ProgramName & "Window Last Width", Me.Width)
            Me.SettingsLocal.SaveSetting(Me.ProgramName & "Window Last Height", Me.Height)
            Me.SettingsLocal.SaveSetting(Me.ProgramName & "Window Last State", Me.WindowState)
        End Sub

        Public Overridable Sub ProgressUpdate(ByVal Plugin As Object, ByVal Current As Integer, ByVal Total As Integer, ByVal Message As String)
            RaiseEvent Do_PluginProgressUpdate(Plugin, Current, Total, Message)
        End Sub

    End Class

    Public Class Window
        Inherits Windows.Forms.Form

        Dim _ParentWin As MyCore.Plugins.IHost = Nothing
        Dim _Opener As Object = Nothing
        Dim _ItemType As String = ""

        Public IsLoaded As Boolean = False
        Public HasChanged As Boolean = False
        Public DisableEnterAsTab As Boolean = False

        Public Event AfterLoad(ByVal Window As MyCore.Plugins.Window)
        Public Event OnEvent(ByVal Window As MyCore.Plugins.Window, ByVal Args As MyCore.Plugins.Window.EventInfo)

        Public Property ParentWin() As MyCore.Plugins.IHost
            Get
                Return Me._ParentWin
            End Get
            Set(ByVal value As MyCore.Plugins.IHost)
                Me._ParentWin = value
            End Set
        End Property

        Public Property Opener() As Object
            Get
                Return Me._Opener
            End Get
            Set(ByVal value As Object)
                Me._Opener = value
            End Set
        End Property

        Public Property ItemType() As String
            Get
                Return Me._ItemType
            End Get
            Set(ByVal value As String)
                Me._ItemType = value
            End Set
        End Property

        Public ReadOnly Property Form() As System.Windows.Forms.Form
            Get
                Return Me
            End Get
        End Property

        Public Sub New()
            Me.KeyPreview = True
            AddHandler MyBase.KeyDown, AddressOf Me.EnterAsTab
        End Sub

        Private Sub Window_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
            If Me.ParentWin IsNot Nothing Then
                Me.ParentWin.Database.LeaveConnectionOpen = True
            End If
        End Sub

        Private Sub Window_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
            RaiseEvent AfterLoad(Me)
            Me.IsLoaded = True
            If Me.ParentWin IsNot Nothing Then
                Me.ParentWin.Database.LeaveConnectionOpen = False
            End If
        End Sub

        Public Sub Open(Optional ByVal Dialog As Boolean = False)
            Me.ParentWin.OpenWindow(Me, Dialog)
        End Sub

        Public Sub Divorce()
            Me.Form.MdiParent = Nothing
        End Sub

        Public Sub Marry()
            Me.Form.MdiParent = Me.ParentWin
        End Sub

        Public Sub NewEvent(ByVal Type As EventType, ByVal ItemType As String, ByVal Item As Object, ByVal Arg As String, Optional ByVal Item2 As Object = Nothing, Optional ByVal Arg2 As String = "")
            RaiseEvent OnEvent(Me, New MyCore.Plugins.Window.EventInfo(Me.Opener, Type, ItemType, Item, Arg, Item2, Arg2))
        End Sub

        Private Sub EnterAsTab(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            If Not Me.DisableEnterAsTab Then
                If e.KeyCode = Keys.Enter Then
                    Dim Form As Windows.Forms.Form = sender
                    Dim Cancel As Boolean = False
                    Dim Control As Windows.Forms.Control = Form.ActiveControl
                    Dim Type As String = Control.GetType.ToString
                    Dim Parent As String = Control.Parent.GetType.ToString
                    If Type.Contains("Grid") Or Type.Contains("Button") Or Type.Contains("Cell") Then
                        Cancel = True
                    ElseIf Parent.Contains("Grid") Or Parent.Contains("Cell") Then
                        Cancel = True
                    ElseIf Type.Contains("TextBox") Then
                        Dim TextBox As Windows.Forms.TextBox = Control
                        If (TextBox.AcceptsReturn Or TextBox.Multiline) And e.KeyCode = Keys.Enter Then
                            Cancel = True
                        ElseIf TextBox.AcceptsTab And e.KeyCode = Keys.Tab Then
                            Cancel = True
                        End If
                    ElseIf Type.Contains("UltraFormattedTextEditor") Then
                        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab Then
                            Cancel = True
                        End If
                    ElseIf Type.Contains("UltraTextEditor") Then
                        Dim TextBox As Infragistics.Win.UltraWinEditors.UltraTextEditor = Control
                        If (TextBox.AcceptsReturn Or TextBox.Multiline) And e.KeyCode = Keys.Enter Then
                            Cancel = True
                        ElseIf TextBox.AcceptsTab And e.KeyCode = Keys.Tab Then
                            Cancel = True
                        End If
                    End If
                    If Not Cancel Then
                        If e.Modifiers = Keys.Control Then
                            Me.FocusPrev()
                        Else
                            Me.FocusNext()
                        End If
                        e.Handled = True
                    End If
                End If
            End If
        End Sub

        Public Sub FocusNext(Optional ByVal Control As Object = Nothing)
            If Control Is Nothing Then
                Me.SelectNextControl(Me.ActiveControl, True, True, True, True)
            Else
                Me.SelectNextControl(Control, True, True, True, True)
            End If
        End Sub

        Public Sub FocusPrev(Optional ByVal Control As Object = Nothing)
            If Control Is Nothing Then
                Me.SelectNextControl(Me.ActiveControl, False, True, True, True)
            Else
                Me.SelectNextControl(Control, False, True, True, True)
            End If
        End Sub

        Delegate Sub ShowDialogCallback()

        Public Shadows Sub ShowDialog()
            If InvokeRequired Then
                Invoke(New ShowDialogCallback(AddressOf Me.ShowDialog))
            Else
                Me.Form.ShowDialog()
            End If
        End Sub

        Delegate Sub ShowCallback()

        Public Shadows Sub Show()
            If InvokeRequired Then
                Try
                    Invoke(New ShowCallback(AddressOf Me.Show))
                Catch ex As Exception
                    MessageBox.Show("Error opening window. " & ex.toString)
                End Try
            Else
                Me.Form.Show()
            End If
        End Sub

        Delegate Sub OnChooseTemplateID(ByVal Id As Integer)
        Delegate Sub OnChooseTemplateHtml(ByVal Html As String)

        Public Overridable Sub ChooseTemplate(ByVal OnChooseCallback As OnChooseTemplateHtml)
            Dim Browse As New Gravity.ChooseTemplateDialog(Me.ParentWin.Database, Gravity.ChooseTemplateDialog.Mode.OpenTemplate)
            If Browse.ShowDialog = Gravity.Response.OK Then
                OnChooseCallback(Browse.SelectedTemplateHtml)
            End If
        End Sub

        Public Overridable Sub ChooseTemplate(ByVal OnChooseCallback As OnChooseTemplateID)
            Dim Browse As New Gravity.ChooseTemplateDialog(Me.ParentWin.Database, Gravity.ChooseTemplateDialog.Mode.OpenTemplate)
            If Browse.ShowDialog = Gravity.Response.OK Then
                OnChooseCallback(Browse.SelectedTemplateID)
            End If
        End Sub

        Public Overridable Function GetDefaultTemplate(ByVal Name As String) As String
            Dim CertTemplate As Integer = Me.ParentWin.Database.GetOne("SELECT value FROM settings WHERE property='" & Name & "'")
            Dim DefaultCert As String = Me.ParentWin.Database.GetOne("SELECT html FROM template WHERE id=" & CertTemplate)
            Return DefaultCert
        End Function

        Class EventInfo

            Public WhatHappened As MyCore.Plugins.EventType
            Public Opener As Object
            Public Type As String
            Public Item1 As Object
            Public Item1Arg As String
            Public Item2 As Object
            Public Item2Arg As String
            Public CallId As String

            Public Sub New(ByRef Opener As Object, ByVal EventType As EventType, ByVal Type As String)
                Me.WhatHappened = EventType
                Me.Type = Type
                Me.Opener = Opener
            End Sub

            Public Sub New(ByRef Opener As Object, ByVal EventType As EventType, ByVal Type As String, ByVal CallId As String)
                Me.WhatHappened = EventType
                Me.Type = Type
                Me.Opener = Opener
                Me.CallId = CallId
            End Sub

            Public Sub New(ByRef Opener As Object, ByVal EventType As EventType, ByVal Type As String, ByVal Item As Object, ByVal Arg As String, Optional ByVal Item2 As Object = Nothing, Optional ByVal Arg2 As String = "")
                Me.WhatHappened = EventType
                Me.Type = Type
                Me.Item1 = Item
                Me.Item1Arg = Arg
                Me.Item2 = Item2
                Me.Item2Arg = Arg2
                Me.Opener = Opener
            End Sub

        End Class

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class

    Public Class HistoryList

        Dim History As HistoryItem()
        Public MaxItems As Integer = 10

        Public Event NewItemAdded(ByVal Item As HistoryItem, ByVal Index As Integer)
        Public Event ItemsCleared()
        Public Event ItemRemoved(ByVal Item As HistoryItem, ByVal FormerIndex As Integer)

        Public ReadOnly Property Count() As Integer
            Get
                Dim Num As Integer = 0
                Try
                    Num = Me.History.Length
                Catch ex As Exception
                    Num = 0
                End Try
                Return Num
            End Get
        End Property

        Public ReadOnly Property IsEmpty() As Boolean
            Get
                Try
                    If Me.History IsNot Nothing Then
                        If Me.History.Length = 0 Then
                            Return True
                        Else
                            Return False
                        End If
                    Else
                        Return True
                    End If

                Catch
                    Return True
                End Try
            End Get
        End Property

        Public ReadOnly Property IsFull() As Boolean
            Get
                Try
                    If Me.History.Length = Me.MaxItems Then
                        Return True
                    Else
                        Return False
                    End If
                Catch
                    Return False
                End Try
            End Get
        End Property

        Public Sub New(Optional ByVal Max As Integer = 20)
            Me.MaxItems = Max
        End Sub

        Public Sub Add(ByVal ItemType As String, ByVal Id As String, ByVal Display As String)
            Dim Item As New HistoryItem(ItemType, Id, Display)
            Dim Index As Integer = Me.MaxItems - 1
            If Not Me.Contains(Item) Then
                If Me.IsFull Then
                    Dim NewHistory(Me.MaxItems - 1) As HistoryItem
                    For i As Integer = 1 To (Me.MaxItems - 1)
                        NewHistory(i - 1) = Me.History(i)
                    Next
                    Me.History = NewHistory
                Else
                    Index = Me.Count
                    ReDim Preserve Me.History(Index)
                End If
                Me.History(Index) = Item
                RaiseEvent NewItemAdded(Item, Index)
            Else
                ' Move Item to the top
            End If
        End Sub

        Public Sub Clear()
            Dim ClearedItems As HistoryItem()
            Me.History = ClearedItems
            RaiseEvent ItemsCleared()
        End Sub

        Public Function Contains(ByVal NewItem As HistoryItem) As Boolean
            If Not Me.IsEmpty Then
                For Each Item As HistoryItem In Me.History
                    If Item.Id = NewItem.Id And Item.ItemType = NewItem.ItemType Then
                        Return True
                    End If
                Next
            End If
            Return False
        End Function

        Public Function Items(ByVal i As Integer) As HistoryItem
            Return Me.History(i)
        End Function

        Public Function Items() As HistoryItem()
            Return Me.History
        End Function

        Public Function Serialize() As String
            If Not Me.IsEmpty Then
                Dim Out As String = ""
                For Each Item As HistoryItem In Me.History
                    Out &= Item.ItemType.Trim & "|" & Item.Id.Trim & "|" & Item.DisplayName.Trim & ControlChars.CrLf
                Next
                Return Out
            Else
                Return ""
            End If
        End Function

        Public Sub DeSerialize(ByVal Str As String)
            Str = Str.Trim
            If Str.Length > 0 Then
                Dim Lines As String() = Str.Split(ControlChars.CrLf)
                For Each Line As String In Lines
                    Line = Line.Trim
                    Dim Cells As String() = Line.Split("|")
                    If Cells.Length = 3 Then
                        Me.Add(Cells(0).Trim, Cells(1).Trim, Cells(2).Trim)
                    End If
                Next
            End If
        End Sub

    End Class

    Public Class HistoryItem

        Public ItemType As String = ""
        Public DisplayName As String = ""
        Public Id As String = ""

        Public Sub New(ByVal Type As String, ByVal Id As String, ByVal Display As String)
            Me.ItemType = Type
            Me.Id = Id
            Me.DisplayName = Display
        End Sub

        Public Sub New(ByVal Str As String)
            Dim Arr As String() = Str.Split("|")
            Me.ItemType = Arr(0)
            Me.Id = Arr(1)
            Me.DisplayName = Arr(2)
        End Sub

        Public Overrides Function ToString() As String
            Return Me.ItemType & "|" & Me.Id & "|" & Me.DisplayName
        End Function

    End Class

End Namespace


