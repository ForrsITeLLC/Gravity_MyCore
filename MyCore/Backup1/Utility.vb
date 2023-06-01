Imports System.Runtime.InteropServices

Public Class App
    Public Shared ReadOnly Property ApplicationData() As String
        Get
            Dim Dir As System.IO.DirectoryInfo = My.Computer.FileSystem.GetDirectoryInfo(My.Computer.FileSystem.SpecialDirectories.AllUsersApplicationData)
            Dim Path As String = Dir.Parent.Parent.FullName & "\"
            If Not My.Computer.FileSystem.DirectoryExists(Path) Then
                My.Computer.FileSystem.CreateDirectory(Path)
            End If
            Return Path
        End Get
    End Property
End Class

Namespace Utility

    Public Enum FileAction
        Open = 1
        Print = 2
        Email = 3
        View = 4
        Fax = 5
        PrintPreview = 6
        Send = 7
    End Enum

    Public Class Hash

        Public Shared Function MD5(ByVal Input As String) As String
            Dim Ue As New System.Text.UnicodeEncoding
            'Retrieve a byte array based on the source text
            Dim ByteSourceText() As Byte = Ue.GetBytes(Input)
            'Instantiate an MD5 Provider object
            Dim Md5hash As New System.Security.Cryptography.MD5CryptoServiceProvider()
            'Compute the hash value from the source
            Dim ByteHash() As Byte = Md5hash.ComputeHash(ByteSourceText)
            'And convert it to String format for return
            Return System.Convert.ToBase64String(ByteHash)
        End Function

    End Class

    Public Class SimpleEncryption

        Dim _Version As Integer = 1

        Public Property Version() As Integer
            Get
                Return Me._Version
            End Get
            Set(ByVal value As Integer)
                Me._Version = value
            End Set
        End Property

        Public Sub New(ByVal Version As Integer)
            Me._Version = Version
        End Sub

        Public Function Encrypt(ByVal Input As String) As String
            Dim Output As String = ""
            If Me._Version = 1 Or Me._Version = 2 Then
                For i As Integer = 0 To Input.Length - 1
                    Dim Asci As Integer = Asc(Input.Substring(i, 1))
                    Output &= Chr(Asci + Me.GetOffset(Asci, "ENCODE"))
                Next
            End If
            Output = Me.HtmlEncode(Output)
            Return Output
        End Function

        Public Function Decrypt(ByVal Input As String) As String
            Dim Output As String = ""
            Input = Me.HtmlDecode(Input)
            If Me._Version = 1 Or Me._Version = 2 Then
                For i As Integer = 0 To Input.Length - 1
                    Dim Asci As Integer = Asc(Input.Substring(i, 1))
                    Output &= Chr(Asci - Me.GetOffset(Asci, "DECODE"))
                Next
            End If
            Return Output
        End Function

        Private Function GetOffset(ByVal Asci As Integer, ByVal Func As String) As Integer
            If Me._Version = 1 Then
                If Asci Mod 2 = 0 Then
                    ' Even
                    If Func = "ENCODE" Then
                        Return 3
                    Else
                        Return 5
                    End If
                Else
                    ' Odd
                    If Func = "ENCODE" Then
                        Return 5
                    Else
                        Return 3
                    End If
                End If
            ElseIf Me._Version = 2 Then
                If Asci Mod 2 = 0 Then
                    ' Even
                    If Func = "ENCODE" Then
                        Return 5
                    Else
                        Return 3
                    End If
                Else
                    ' Odd
                    If Func = "ENCODE" Then
                        Return 3
                    Else
                        Return 5
                    End If
                End If
            End If
        End Function

        Private Function HtmlEncode(ByVal Input As String) As String
            Dim Output As String = Input
            Output = Output.Replace("&", "&amp;")
            Output = Output.Replace(">", "&gt;")
            Output = Output.Replace("<", "&lt;")
            Return Output
        End Function

        Private Function HtmlDecode(ByVal Input As String) As String
            Dim Output As String = Input
            Output = Output.Replace("&amp;", "&")
            Output = Output.Replace("&gt;", ">")
            Output = Output.Replace("&lt;", "<")
            Return Output
        End Function

    End Class

    Public Class PrintTool

        Dim _File As String = ""

        Public Sub New(ByVal FilePath As String)
            Me._File = FilePath
        End Sub

        Public Sub Print()
            Dim Print As New System.Diagnostics.ProcessStartInfo
            Print.FileName = Me._File
            Print.Verb = "print"
            Print.WindowStyle = ProcessWindowStyle.Hidden
            Print.UseShellExecute = True
            Try
                System.Diagnostics.Process.Start(Print)
            Catch ex As Exception
                Throw New Exception(Print.ToString & " -- " & ex.ToString)
            End Try
        End Sub
    End Class

    Public Class CsvFile

        Public Shared Sub FromDataTable(ByVal dt As DataTable, ByVal FileName As String, Optional ByVal Delimiter As String = ControlChars.Tab)
            Dim Out As String = ""
            If dt IsNot Nothing Then
                For Each c As DataColumn In dt.Columns
                    Out &= c.ColumnName & Delimiter
                Next
                Out &= ControlChars.CrLf
                For Each r As DataRow In dt.Rows
                    For i As Integer = 0 To dt.Columns.Count - 1
                        If r.Item(i) IsNot DBNull.Value Then
                            Out &= r.Item(i).ToString
                        End If
                        Out &= Delimiter
                    Next
                    Out &= ControlChars.CrLf
                Next
            End If
            My.Computer.FileSystem.WriteAllText(FileName, Out, False, System.Text.Encoding.Default)
        End Sub

    End Class

    Public Class Web

        Shared Function DownloadFile(ByVal uri As String, ByVal destFile As String, _
        Optional ByVal username As String = Nothing, Optional ByVal pwd As String = _
        Nothing) As Boolean
            Dim wc As New System.Net.WebClient
            Dim myWebClient As New System.Net.WebClient
            If Not username Is Nothing AndAlso Not pwd Is Nothing Then
                wc.Credentials = New System.Net.NetworkCredential(username, pwd)
            End If

            Try
                myWebClient.DownloadFile(uri, destFile)
                Return True
            Catch ex As Exception
                Return False
            End Try

        End Function

    End Class

    Public Class Downloader

        Public Url As String = ""
        Private IsDownloading As Boolean = False

        Public Sub New(ByVal url As String)
            Me.Url = url
        End Sub

        Public Function Download() As String

            If Not IsDownloading Then
                IsDownloading = True

                Dim Request As Net.WebRequest = Net.WebRequest.Create(Url)
                Dim Response As Net.WebResponse = Request.GetResponse()
                Dim intTotalBuff As Integer = 0
                Dim Buffer(128) As [Byte]
                Dim Stream As IO.Stream = Response.GetResponseStream()
                intTotalBuff = Stream.Read(Buffer, 0, 128)
                Dim strRes As New System.Text.StringBuilder("")

                While intTotalBuff <> 0
                    strRes.Append(System.Text.Encoding.ASCII.GetString(Buffer, 0, intTotalBuff))
                    intTotalBuff = Stream.Read(Buffer, 0, 128)
                End While

                IsDownloading = False

                Return strRes.ToString

            Else
                Throw New Exception("Already downloading")
            End If

        End Function

    End Class


    Public Class Workstation

        Public Shared Identity As System.Security.Principal.WindowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent
        Public Shared Principal As System.Security.Principal.WindowsPrincipal = New System.Security.Principal.WindowsPrincipal(Identity)

        Public Shared Function IpAddress(Optional ByVal Index As Integer = -1) As String
            Dim AddressList As System.Array = System.Net.Dns.GetHostByName(System.Net.Dns.GetHostName()).AddressList
            'Return AddressList(0).ToString
            If Index < 0 Then
                Return AddressList(AddressList.Length - 1).ToString
            Else
                Return AddressList(Index).ToString
            End If
        End Function

        Public Shared Function IpAddresses() As System.Array
            Dim AddressList As System.Array = System.Net.Dns.GetHostByName(System.Net.Dns.GetHostName()).AddressList
            Return AddressList
        End Function

        Shared Function ComputerName() As String
            Return System.Environment.MachineName
        End Function

        Shared Function OSVersion() As String
            Return System.Environment.OSVersion.Version.ToString
        End Function

        Shared Function OSName() As String
            Return System.Environment.OSVersion.Platform.ToString
        End Function

        Shared Function WindowsUser() As String
            Try
                Dim mc As New System.Management.ManagementClass("Win32_Process")
                Dim moc As System.Management.ManagementObjectCollection = mc.GetInstances
                Dim mo As System.Management.ManagementObject

                For Each mo In moc

                    If mo.Item("Name").ToString.Trim = "explorer.exe" Then
                        Dim argList As String() = {String.Empty}
                        Dim objReturn As Object = mo.InvokeMethod("GetOwner", argList)

                        If Convert.ToInt32(objReturn) = 0 Then
                            Dim userName As String = argList(0)
                            Return userName
                            Exit For
                        End If
                    End If

                Next
            Catch
                ' Nothing
            End Try
            Dim User As String = My.User.CurrentPrincipal.Identity.Name
            If User.Contains("\") Then
                User = User.Substring(User.LastIndexOf("\") + 1)
            End If
            If User = "User" Or User = "System" Then
                User = Environment.UserName
            End If
            If User = "User" Or User = "System" Then
                User = System.Security.Principal.WindowsIdentity.GetCurrent.Name
                If User.Contains("\") Then
                    User = User.Substring(User.LastIndexOf("\") + 1)
                End If
            End If
            Return User
        End Function

        Shared Function ProcessorId() As String
            Dim mc As System.Management.ManagementClass
            Dim mo As System.Management.ManagementObject
            Dim id As String = ""
            Try
                mc = New System.Management.ManagementClass("Win32_Processor")
                Dim moc As System.Management.ManagementObjectCollection = mc.GetInstances
                For Each mo In moc
                    Try
                        id = mo.Item("ProcessorId").ToString()
                    Catch ex As Exception
                        id = ""
                    End Try
                Next
            Catch
                id = ""
            End Try
            Return id
        End Function

        Shared Function BiosId() As String
            Dim mc As System.Management.ManagementClass
            Dim mo As System.Management.ManagementObject
            Dim id As String = ""
            Try
                mc = New System.Management.ManagementClass("Win32_BIOS")
                Dim moc As System.Management.ManagementObjectCollection = mc.GetInstances
                For Each mo In moc
                    Try
                        id = mo.Item("SerialNumber").ToString()
                    Catch ex As Exception
                        id = ""
                    End Try
                Next
            Catch
                id = ""
            End Try
            Return id
        End Function

        Shared Function NetworkConnections() As DataTable
            Dim Table As New DataTable
            Dim dr As DataRow
            Dim adapter As DataRow
            Dim i As Integer
            Table.Columns.Add("ID")
            Table.Columns.Add("MAC Address")
            Table.Columns.Add("IP Address")
            Table.Columns.Add("Subnet Mask")
            Table.Columns.Add("Gateway")
            Table.Columns.Add("Description")
            Table.Columns.Add("DHCP Enabled")
            Table.Columns.Add("DHCP Server")
            Table.Columns.Add("DHCP Lease Obtained")
            Table.Columns.Add("DHCP Lease Expires")
            Table.Columns.Add("Connection Name")
            Table.Columns.Add("Connection Status")
            Dim AdaptersConfig As System.Management.ManagementObjectSearcher
            Dim Results As System.Management.ManagementObject
            Dim Item As Object
            Dim Scope As New System.Management.ManagementScope("\\" & Environment.MachineName & "\root\cimv2")
            Dim Query As New System.Management.SelectQuery("SELECT * FROM Win32_NetworkAdapterConfiguration")
            AdaptersConfig = New System.Management.ManagementObjectSearcher(Scope, Query)
            For Each Results In AdaptersConfig.Get
                If Results.Item("IPEnabled") Then
                    dr = Table.NewRow
                    adapter = GetConnection(Results.Item("Index"))
                    dr.Item("ID") = Results.Item("Index")
                    dr.Item("Connection Name") = adapter.Item("Connection Name")
                    dr.Item("Connection Status") = adapter.Item("Connection Status")
                    dr.Item("MAC Address") = Results.Item("MACAddress")
                    For i = 0 To Results.Item("IPAddress").Length - 1
                        dr.Item("IP Address") &= Results.Item("IPAddress")(i) & " "
                    Next
                    For i = 0 To Results.Item("IPSubnet").Length - 1
                        dr.Item("Subnet Mask") &= Results.Item("IPSubnet")(i) & " "
                    Next
                    For i = 0 To Results.Item("DefaultIPGateway").Length - 1
                        dr.Item("Gateway") &= Results.Item("DefaultIPGateway")(i) & " "
                    Next
                    dr.Item("Description") = Results.Item("Description")
                    dr.Item("DHCP Enabled") = Results.Item("DHCPEnabled")
                    dr.Item("DHCP Server") = Results.Item("DHCPServer")
                    Table.Rows.Add(dr)
                End If
            Next
            Return Table
        End Function

        Public Shared Function GetConnection(ByVal i As UInt32) As DataRow
            Dim Adapters As System.Management.ManagementObjectSearcher
            Dim Scope As New System.Management.ManagementScope("\\" & Environment.MachineName & "\root\cimv2")
            Dim Query As New System.Management.SelectQuery("SELECT * FROM Win32_NetworkAdapter WHERE Index=" & i.ToString)
            Adapters = New System.Management.ManagementObjectSearcher(Scope, Query)
            Dim Results As System.Management.ManagementObject
            Dim dt As New DataTable
            dt.Columns.Add("Connection Name")
            dt.Columns.Add("Connection Status")
            Dim dr As DataRow = dt.NewRow
            For Each Results In Adapters.Get
                dr.Item("Connection Name") = Results.Item("NetConnectionId")
                dr.Item("Connection Status") = Results.Item("NetConnectionStatus")
                Exit For
            Next
            Return dr
        End Function

        Shared Function GetRealName() As String
            'Return System.Environment
        End Function

        Shared Function Domain() As String
            Return Identity.Name.Substring(0, Identity.Name.IndexOf("\"))
        End Function

        Shared Function isUserInGroup(ByVal strGroupName As String) As Boolean
            Return Principal.IsInRole(strGroupName)
        End Function

        Public Class NetworkAdapter
            Dim objCS As System.Management.ManagementObjectSearcher
            Dim objMgmt As System.Management.ManagementObject

            Public Sub New()
                Dim obj As Object
                Dim Scope As New System.Management.ManagementScope("\\" & Environment.MachineName & "\root\cimv2")
                Dim Query As New System.Management.SelectQuery("SELECT * FROM Win32_NetworkAdapterConfiguration")
                objCS = New System.Management.ManagementObjectSearcher(Scope, Query)
                For Each objMgmt In objCS.Get

                Next
            End Sub

        End Class

        Public Class OperatingSystem

            Private Const SM_REMOTESESSION As Integer = &H1000

            Dim _BootDevice As String
            Dim _Caption As String
            Dim _Version As String
            Dim _ServicePack As String
            Dim _InstallDate As String
            Dim _SerialNumber As String

            Private Declare Function GetSystemMetrics Lib "User32" (ByVal nIndex As Integer) As Integer

            Public Shared Function IsRemoteDesktop() As Boolean
                If GetSystemMetrics(SM_REMOTESESSION) <> 0 Then
                    Return True
                Else
                    Return False
                End If
            End Function

            Public ReadOnly Property BootDevice()
                Get
                    Return _BootDevice
                End Get
            End Property

            Public ReadOnly Property Name()
                Get
                    Return _Caption
                End Get
            End Property

            Public ReadOnly Property Version()
                Get
                    Return _Version
                End Get
            End Property

            Public ReadOnly Property ServicePack()
                Get
                    Return _ServicePack
                End Get
            End Property

            Public ReadOnly Property InstallDate()
                Get
                    Return _InstallDate
                End Get
            End Property

            Public ReadOnly Property SerialNumber()
                Get
                    Return _SerialNumber
                End Get
            End Property


            Public Sub New()
                Dim objOS As System.Management.ManagementObjectSearcher
                Dim objResults As System.Management.ManagementObject
                Dim objItem As Object
                Dim Scope As New System.Management.ManagementScope("\\" & Environment.MachineName & "\root\cimv2")
                Dim Query As New System.Management.SelectQuery("SELECT * FROM Win32_OperatingSystem")
                objOS = New System.Management.ManagementObjectSearcher(Scope, Query)
                For Each objResults In objOS.Get
                    _BootDevice = objResults("BootDevice").ToString
                    _Caption = objResults("Caption").ToString
                    _Version = objResults("Version").ToString
                    _ServicePack = objResults("CSDVersion").ToString
                    _InstallDate = objResults("InstallDate").ToString
                    _SerialNumber = objResults("SerialNumber").ToString
                Next
            End Sub

        End Class

        Public Class Computer
            Dim objCS As System.Management.ManagementObjectSearcher
            Dim objMgmt As System.Management.ManagementObject
            Dim _strManufacturer As String
            Dim _strModel As String
            Dim _strSysType As String
            Dim _strMemory As String

            Public ReadOnly Property Manufacturer()
                Get
                    Return _strManufacturer
                End Get
            End Property

            Public ReadOnly Property Model()
                Get
                    Return _strModel
                End Get
            End Property

            Public ReadOnly Property SysType()
                Get
                    Return _strSysType
                End Get
            End Property

            Public ReadOnly Property TotalPhysicalMemory()
                Get
                    Return _strMemory
                End Get
            End Property

            Public Sub New()
                Dim obj As Object
                Dim Scope As New System.Management.ManagementScope("\\" & Environment.MachineName & "\root\cimv2")
                Dim Query As New System.Management.SelectQuery("SELECT * FROM Win32_ComputerSystem")
                objCS = New System.Management.ManagementObjectSearcher(Scope, Query)
                For Each objMgmt In objCS.Get
                    _strManufacturer = objMgmt("manufacturer").ToString
                    _strModel = objMgmt("model").ToString
                    _strSysType = objMgmt("systemtype").ToString
                    _strMemory = objMgmt("totalphysicalmemory").ToString
                Next
            End Sub

        End Class


    End Class


    Public Class Idle

        'THIS CLASS WAS TAKEN FROM A MAILING LIST MESSAGE AVAILABLE AT 
        'http://groups.google.de/group/microsoft.public.de.german.entwickler.dotnet.vb/msg/89371099848cfd7a 
        'AND IS ATTRIBUTED (AS FAR AS I CAN TELL) TO Diana Mueller AS THE MESSAGE POSTER 
        'BECAUSE IT WAS POSTED TO A PUBLIC MAILING LIST I CONSIDER IT TO BE IN THE PUBLIC DOMAIN. 

        'THE CODE HOOKS INTO A LASTINPUTINFO DLL CALL PROVIDING THE IDLE TIME (IN TICKS) OF THE USER 

        'INFO ON THE DLL CALL CAN BE FOUND AT http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/userinput/keyboardinput/keyboardinputreference/keyboardinputfunctions/getlastinputinfo.asp 
        'AKA http://tinyurl.com/3ul2r 

        'NOTE THAT THE RETURN VALUE OF IdleTimeInTicks() DIVIDED BY 1000 IS IDLE TIME IN SECONDS 

        'Mike Hillyer - July 18 2005 

        Private Declare Function GetLastInputInfo Lib "User32.dll" (ByRef lii As LASTINPUTINFO) As Boolean

        <StructLayout(LayoutKind.Sequential)> Public Structure LASTINPUTINFO

            Public cbSize As Int32

            Public dwTime As Int32

        End Structure

        Public Shared ReadOnly Property IdleTimeInTicks() As Int32

            Get

                Dim lii As New LASTINPUTINFO

                lii.cbSize = Marshal.SizeOf(lii)

                If GetLastInputInfo(lii) Then

                    Return Environment.TickCount - lii.dwTime

                End If

            End Get

        End Property

    End Class


    Public Class ListViewComparer
        Implements IComparer

        Private m_ColumnNumber As Integer
        Private m_SortOrder As System.Windows.Forms.SortOrder

        Public Sub New(ByVal column_number As Integer, ByVal sort_order As System.Windows.Forms.SortOrder)
            m_ColumnNumber = column_number
            m_SortOrder = sort_order
        End Sub

        ' Compare the items in the appropriate column
        ' for objects x and y.
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            Dim item_x As System.Windows.Forms.ListViewItem = DirectCast(x, System.Windows.Forms.ListViewItem)
            Dim item_y As System.Windows.Forms.ListViewItem = DirectCast(y, System.Windows.Forms.ListViewItem)

            ' Get the sub-item values.
            Dim string_x As String
            If item_x.SubItems.Count <= m_ColumnNumber Then
                string_x = ""
            Else
                string_x = item_x.SubItems(m_ColumnNumber).Text
            End If

            Dim string_y As String
            If item_y.SubItems.Count <= m_ColumnNumber Then
                string_y = ""
            Else
                string_y = item_y.SubItems(m_ColumnNumber).Text
            End If

            ' Compare them.
            If m_SortOrder = System.Windows.Forms.SortOrder.Ascending Then
                Return String.Compare(string_x, string_y)
            Else
                Return String.Compare(string_y, string_x)
            End If
        End Function
    End Class

    Public Class Tools


        Public Shared Sub ExportToExcel(ByVal dt As DataTable, ByVal arrColumnNames As Array)

            ' Integer Variables
            Dim intRow As Integer
            Dim intCol As Integer

            ' Excel Variables
            Dim exlApp As Object = CreateObject("Excel.Application")
            Dim exlBook As Object = exlApp.Workbooks.Add
            Dim exlSheet As Object = exlBook.Worksheets(1)

            ' Hide Excel
            exlApp.Visible = False

            For intCol = 1 To arrColumnNames.Length
                Application.DoEvents()
                exlSheet.Cells(1, intCol).Value = arrColumnNames(intCol - 1)
                Application.DoEvents()
            Next

            ' Write Data
            For intRow = 1 To dt.Rows.Count
                For intCol = 1 To dt.Columns.Count
                    exlSheet.Cells(intRow + 1, intCol).Value = dt.Rows(intRow - 1).Item(intCol - 1)
                Next
            Next

            'Show Excel
            exlApp.Visible = True

        End Sub

        Public Shared Sub Print(ByRef Browser As AxSHDocVw.AxWebBrowser)
            MyCore.Utility.Tools.Print(Browser.LocationURL, Browser)
        End Sub

        Public Shared Sub Print(ByVal FilePath As String, ByRef Browser As AxSHDocVw.AxWebBrowser)
            ' Open File
            Browser.Navigate(FilePath)
            While Browser.QueryStatusWB(SHDocVw.OLECMDID.OLECMDID_PRINT) <> SHDocVw.OLECMDF.OLECMDF_SUPPORTED + SHDocVw.OLECMDF.OLECMDF_ENABLED
                Application.DoEvents()
            End While
            ' Change registry
            Dim regkey As Microsoft.Win32.RegistryKey
            Dim strKey As String = "Software\\Microsoft\\Internet Explorer\\PageSetup"
            regkey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(strKey, True)
            'save the current header/footer strings
            Dim OldHeader As String = ""
            Dim OldFooter As String = ""
            Dim ChangedKey As Boolean = False
            Try
                OldHeader = regkey.GetValue("header")
                OldFooter = regkey.GetValue("footer")
                'set our new header/footer strings
                regkey.SetValue("header", String.Empty)
                regkey.SetValue("footer", String.Empty)
                regkey.Close()
                ChangedKey = True
            Catch ex As Exception
                Dim Err As New MyCore.Gravity.ErrorBox("Failed to edit Windows Registry.  This might be a permission issue. Press continue, this is not a serious error.", "Change RegKey Error", ex.ToString, 1)
            End Try
            ' Print
            Browser.ExecWB(SHDocVw.OLECMDID.OLECMDID_PRINT, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_PROMPTUSER)
            ' Change registry back
            If ChangedKey Then
                Try
                    regkey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(strKey, True)
                    regkey.SetValue("header", OldHeader)
                    regkey.SetValue("footer", OldFooter)
                Catch ex As Exception
                    Dim Err As New MyCore.Gravity.ErrorBox("Failed to edit Windows Registry.  This might be a permission issue. Press continue, this is not a serious error.", "Change RegKey Error", ex.ToString, 1)
                End Try
            End If
        End Sub

    End Class


    Public Class Terminology

        Public Shared Function MonthFrequency(ByVal intMonths As Integer) As String
            Select Case intMonths
                Case 1
                    Return "Monthly"
                Case 2
                    Return "Semi-Monthly"
                Case 3
                    Return "Quarterly"
                Case 4
                    Return "Tri-Annual"
                Case 6
                    Return "Semi-Annual"
                Case 12
                    Return "Annual"
                Case Else
                    Return ""
            End Select
        End Function

    End Class


    Public Class Calendar

        Public Shared Function LastDayOfMonth(ByVal dDate As Date) As Date
            Dim dNextMonth As New Date(dDate.AddMonths(1).Year, dDate.AddMonths(1).Month, 1)
            Return dNextMonth.AddDays(-1)
        End Function


        Public Sub New()

        End Sub

    End Class

    Public Class Communication

        Dim _SettingsGlobal As MyCore.cSettings
        Dim CurrentUser As MyCore.cEmployee

        Public Sub New(ByVal Settings As MyCore.cSettings, ByVal User As MyCore.cEmployee)
            Me._SettingsGlobal = Settings
            Me.CurrentUser = User
        End Sub

        Public Function CreateNewEmail(ByVal ToAddress As String, ByVal ToName As String) As MyCore.Email.Message
            Dim Settings As New MyCore.Email.MailSettings
            Settings.Host = Me._SettingsGlobal.GetValue("SMTP Host")
            Settings.UserName = Me._SettingsGlobal.GetValue("SMTP User")
            Settings.Password = Me._SettingsGlobal.GetValue("SMTP Password")
            Settings.Port = Me._SettingsGlobal.GetValue("SMTP Port")
            Settings.Domain = Me._SettingsGlobal.GetValue("SMTP Domain")
            Settings.EnableSSL = Me._SettingsGlobal.GetValue("SMTP Enable SSL")
            Dim Message As New MyCore.Email.Message(Settings)
            Message.AddToAddress(ToAddress, ToName)
            Message.FromAddress = Me.CurrentUser.Email
            Message.FromName = Me.CurrentUser.LastName & ", " & Me.CurrentUser.FirstName
            Message.ReplyToAddress = Me.CurrentUser.Email
            Message.ReplyToName = Me.CurrentUser.LastName & ", " & Me.CurrentUser.FirstName
            Return Message
        End Function

        Public Sub SendFax(ByVal FilePath As String, ByVal FaxNumber As String, ByVal FaxName As String)
            Dim FaxServerPath As String = Me._SettingsGlobal.GetValue("Fax Server")
            If FaxServerPath.Length > 0 Then
                Dim Server As New FAXCOMLib.FaxServer
                Dim Document As FAXCOMLib.FaxDoc
                Server.Connect(FaxServerPath)
                Document = Server.CreateDocument(FilePath)
                Document.FaxNumber = FaxNumber
                Document.RecipientName = FaxName
                Document.DisplayName = FaxName
                Document.Send()
                Server.Disconnect()
            Else
                Throw New Exception("No fax server set.")
            End If
        End Sub

    End Class

End Namespace