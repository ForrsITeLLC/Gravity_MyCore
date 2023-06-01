Imports System.Xml
Imports System
Imports System.Security.Cryptography

Namespace Settings

    Public Class ConfigFile

        Public File As String = ""

        Public ReadOnly Property EncVersion() As Double
            Get
                Dim Doc As New XmlDocument
                Doc.Load(Me.File)
                Return Doc.SelectSingleNode("/gravity/@enc").InnerText
            End Get
        End Property

        Public ReadOnly Property SchemaVersion() As Double
            Get
                Dim Doc As New XmlDocument
                Doc.Load(Me.File)
                Return Doc.SelectSingleNode("/gravity/@schema").InnerText
            End Get
        End Property

        Public Sub New(ByVal FilePath As String)
            Me.File = FilePath
            InitializeConfigFile()
        End Sub

        'initialize the apps config file, create it if it doesn't exist
        Private Sub InitializeConfigFile()

            If Not My.Computer.FileSystem.FileExists(Me.File) Then
                Dim Contents As String = ""
                Contents &= "<?xml version=""1.0"" encoding=""utf-8""?>" & ControlChars.CrLf
                Contents &= "<gravity schema=""1.0"" enc=""1"">" & ControlChars.CrLf
                Contents &= "<settings>" & ControlChars.CrLf
                Contents &= "</settings>" & ControlChars.CrLf
                Contents &= "<translation lang=""en-us"">" & ControlChars.CrLf
                Contents &= "</translation>" & ControlChars.CrLf
                Contents &= "</gravity>" & ControlChars.CrLf
                My.Computer.FileSystem.WriteAllText(Me.File, Contents, False)
            End If

        End Sub

        'get an application setting by key value
        Public Function GetSetting(ByVal Key As String, Optional ByVal DefaultVal As String = "") As String
            ' Load XML File
            Dim Doc As New XmlDocument
            Doc.Load(Me.File)
            ' Look for key
            Dim Node As XmlNode = Doc.SelectSingleNode("/gravity/settings/key[@name='" & Key & "']")
            ' If Node Exists, Return It
            If Node IsNot Nothing Then
                Return Node.InnerText
            Else
                ' Otherwise Create it
                Me.AddSetting(Doc, Key, DefaultVal)
                Return DefaultVal
            End If
        End Function

        Public Function GetTerm(ByVal Term As String, Optional ByVal Form As String = "singular", Optional ByVal Lang As String = "en-us") As String
            ' Load XML File
            Dim Doc As New XmlDocument
            Doc.Load(Me.File)
            ' Look for key
            Dim Node As XmlNode = Doc.SelectSingleNode("/gravity/translation[@lang='" & Lang & "']/key[@name='" & Term & "']/" & Form)
            ' If Node Exists, Return It
            If Node IsNot Nothing Then
                Return Node.InnerText
            Else
                Me.SaveTerm(Doc, Term, Form, Term, Lang)
                Return Term
            End If
        End Function

        'save an application setting, takes a key and a value
        Public Sub SaveSetting(ByVal key As String, ByVal value As String)
            ' Load XML File
            Dim Doc As New XmlDocument
            Doc.Load(Me.File)
            ' Look for key
            Dim Node As XmlNode = Doc.SelectSingleNode("/gravity/settings/key[@name='" & key & "']")
            ' If Node Exists, Return It
            If Node IsNot Nothing Then
                Node.InnerText = value
                Doc.Save(Me.File)
            Else
                Me.AddSetting(Doc, key, value)
            End If
        End Sub

        Private Sub AddSetting(ByRef Doc As System.Xml.XmlDocument, ByVal key As String, Optional ByVal value As String = "")
            Dim xmlParent As Xml.XmlNode = Doc.SelectSingleNode("/gravity/settings")
            Dim xmlElement As Xml.XmlElement = Doc.CreateElement("key")
            Dim xmlKey As Xml.XmlAttribute = Doc.CreateAttribute("name")
            xmlKey.Value = key
            xmlElement.InnerText = value
            xmlElement.Attributes.Append(xmlKey)
            xmlParent.AppendChild(xmlElement)
            Doc.Save(Me.File)
        End Sub

        Public Sub SaveTerm(ByVal Term As String, ByVal Form As String, ByVal Custom As String, Optional ByVal Lang As String = "en")
            Dim Doc As New XmlDocument
            Doc.Load(Me.File)
            Me.SaveTerm(Doc, Term, Form, Custom, Lang)
        End Sub


        Private Sub SaveTerm(ByRef Doc As System.Xml.XmlDocument, ByVal Term As String, ByVal Form As String, ByVal Custom As String, ByVal Lang As String)
            Dim Parent As XmlNode = Doc.SelectSingleNode("/gravity/translation[@lang='" & Lang & "']")
            Dim xmlTerm As XmlNode = Parent.SelectSingleNode("key[@name='" & Term & "']")
            Dim xmlForm As XmlNode = Parent.SelectSingleNode("key[@name='" & Term & "']/" & Form)
            ' If term doesn't exist, create it
            If xmlTerm Is Nothing Then
                xmlTerm = Doc.CreateElement("key")
                xmlTerm.Attributes.Append(Doc.CreateAttribute("name"))
                xmlTerm.InnerText = Term
                Parent.AppendChild(xmlTerm)
            End If
            ' If form doesn't exit
            If xmlForm Is Nothing Then
                ' Now Create the form element
                xmlForm = Doc.CreateElement(Form)
                xmlTerm.AppendChild(xmlForm)
            End If
            ' Set value
            xmlForm.InnerText = Custom
            ' Save changes
            Doc.Save(Me.File)
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

    End Class

    Public Class AppSettings

        Public ReadOnly Property Folder() As String
            Get
                'Return My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\Evware\"
                Return My.Computer.FileSystem.CurrentDirectory & "\"
            End Get
        End Property

        Public ReadOnly Property AppFolder() As String
            Get
                'Return Me.Folder & "\" & Me.AppName & "\"
                Return Me.Folder
            End Get
        End Property

        Public ReadOnly Property FilePath() As String
            Get
                Return Me.Folder & Me.AppName & ".xml"
            End Get
        End Property

        Public ReadOnly Property IsOK() As Boolean
            Get
                Dim Doc As New XmlDocument
                Try
                    Doc.Load(Me.FilePath)
                    Return True
                Catch
                    Return False
                End Try
            End Get
        End Property

        Dim AppName As String = "Globals"

        Public Sub New(Optional ByVal strAppName As String = "Globals")
            ' set name
            Me.AppName = strAppName
            ' Make sure folder exists
            If Not My.Computer.FileSystem.DirectoryExists(Me.Folder) Then
                My.Computer.FileSystem.CreateDirectory(Me.Folder)
            End If
            If Not My.Computer.FileSystem.DirectoryExists(Me.AppFolder) Then
                My.Computer.FileSystem.CreateDirectory(Me.AppFolder)
            End If
            ' Create settings file if it doesn't exist
            If Not My.Computer.FileSystem.FileExists(Me.FilePath) Then
                Me.CreateBlankSettingsFile()
            End If
            ' Check that the format is okay
            If Not Me.IsOK Then
                MessageBox.Show("Local settings file was corrupt.  Creating a new blank file.", "Gravity Settings", MessageBoxButtons.OK)
                Me.CreateBlankSettingsFile()
            End If
        End Sub

        Private Sub CreateBlankSettingsFile()
            Dim Contents As String = ""
            Contents &= "<?xml version=""1.0"" encoding=""utf-8""?>" & ControlChars.CrLf
            Contents &= "<configuration>" & ControlChars.CrLf
            Contents &= "  <appSettings>" & ControlChars.CrLf
            Contents &= "  </appSettings>" & ControlChars.CrLf
            Contents &= "  <translation lang=""en"">" & ControlChars.CrLf
            Contents &= "  </translation>" & ControlChars.CrLf
            Contents &= "</configuration>" & ControlChars.CrLf
            My.Computer.FileSystem.WriteAllText(Me.FilePath, Contents, False)
        End Sub

        'get an application setting by key value
        Public Function GetSetting(ByVal Key As String, Optional ByVal DefaultVal As String = "") As String
            ' Load XML File
            Dim Doc As New XmlDocument
            Doc.Load(Me.FilePath)
            ' Look for key
            Dim Node As XmlNode = Doc.SelectSingleNode("/configuration/appSettings/add[@key='" & Key & "']")
            ' If Node Exists, Return It
            If Node IsNot Nothing Then
                Return Node.Attributes.GetNamedItem("value").InnerText
            Else
                ' Otherwise Create it
                Me.AddSetting(Doc, Key, DefaultVal)
                Return DefaultVal
            End If
        End Function

        Public Function GetTerm(ByVal Term As String, Optional ByVal Form As String = "singular", Optional ByVal Lang As String = "en") As String
            ' Load XML File
            Dim Doc As New XmlDocument
            Doc.Load(Me.FilePath)
            ' Look for key
            Dim Node As XmlNode = Doc.SelectSingleNode("/configuration/translation[@lang='" & Lang & "']/term[@key='" & Term & "']/" & Form)
            ' If Node Exists, Return It
            If Node IsNot Nothing Then
                Return Node.InnerText
            Else
                Me.SaveTerm(Doc, Term, Form, Term, Lang)
                Return Term
            End If
        End Function

        'save an application setting, takes a key and a value
        Public Sub SaveSetting(ByVal key As String, ByVal value As String)
            ' Load XML File
            Dim Doc As New XmlDocument
            Doc.Load(Me.FilePath)
            ' Look for key
            Dim Node As XmlNode = Doc.SelectSingleNode("/configuration/appSettings/add[@key='" & key & "']")
            ' If Node Exists, Return It
            If Node IsNot Nothing Then
                Node.Attributes.GetNamedItem("value").InnerText = value
                Doc.Save(Me.FilePath)
            Else
                Me.AddSetting(Doc, key, value)
            End If
        End Sub

        Private Sub AddSetting(ByRef Doc As System.Xml.XmlDocument, ByVal key As String, Optional ByVal value As String = "")
            Dim xmlParent As Xml.XmlNode = Doc.SelectSingleNode("/configuration/appSettings")
            Dim xmlElement As Xml.XmlElement = Doc.CreateElement("add")
            Dim xmlKey As Xml.XmlAttribute = Doc.CreateAttribute("key")
            Dim xmlValue As Xml.XmlAttribute = Doc.CreateAttribute("value")
            xmlKey.InnerText = key
            xmlValue.InnerText = value
            xmlElement.Attributes.Append(xmlKey)
            xmlElement.Attributes.Append(xmlValue)
            xmlParent.AppendChild(xmlElement)
            Doc.Save(Me.FilePath)
        End Sub

        Public Sub SaveTerm(ByVal Term As String, ByVal Form As String, ByVal Custom As String, Optional ByVal Lang As String = "en")
            Dim Doc As New XmlDocument
            Doc.Load(Me.FilePath)
            Me.SaveTerm(Doc, Term, Form, Custom, Lang)
        End Sub

        Private Sub SaveTerm(ByRef Doc As System.Xml.XmlDocument, ByVal Term As String, ByVal Form As String, ByVal Custom As String, ByVal Lang As String)
            Dim Parent As XmlNode = Doc.SelectSingleNode("/configuration/translation[@lang='" & Lang & "']")
            Dim xmlTerm As XmlNode = Parent.SelectSingleNode("term[@key='" & Term & "']")
            Dim xmlForm As XmlNode = Parent.SelectSingleNode("term[@key='" & Term & "']/" & Form)
            ' If term doesn't exist, create it
            If xmlTerm Is Nothing Then
                xmlTerm = Doc.CreateElement("term")
                xmlTerm.Attributes.Append(Doc.CreateAttribute("key"))
                xmlTerm.Attributes.GetNamedItem("key").InnerText = Term
                Parent.AppendChild(xmlTerm)
            End If
            ' If form doesn't exit
            If xmlForm Is Nothing Then
                ' Now Create the form element
                xmlForm = Doc.CreateElement(Form)
                xmlTerm.AppendChild(xmlForm)
            End If
            ' Set value
            xmlForm.InnerText = Custom
            ' Save changes
            Doc.Save(Me.FilePath)
        End Sub

    End Class

End Namespace
