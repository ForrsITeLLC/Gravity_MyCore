Imports System
Imports System.Security.Cryptography

Namespace Gravity

    Public Enum Response
        Yes = 2
        OK = 1
        Cancel = 0
        No = -1
        Quit = -2
    End Enum

    Public Class InfoBox

        Dim _Message As String
        Dim _Title As String
        Dim _Window As New InfoPop
        Dim _Button As Response

        Public ReadOnly Property ButtonPress() As Response
            Get
                Return _Button
            End Get
        End Property

        Public Sub New(ByVal Message As String, Optional ByVal Title As String = "Gravity Info")
            Me.SetValues(Message, Title)
            Me.ShowDialog()
        End Sub

        Public Sub New()
            ' Nada
        End Sub

        Private Sub SetValues(ByVal Message As String, ByVal Title As String)
            Me._Message = Message
            Me._Title = Title
        End Sub

        Public Function ShowDialog(ByVal Message As String, Optional ByVal Title As String = "Gravity Info") As Response
            Me.SetValues(Message, Title)
            Me.ShowDialog()
            Return Me.ButtonPress
        End Function

        Public Function ShowDialog() As Response
            Me._Window.lblCaption.Text = Me._Message
            Me._Window.Text = Me._Title
            AddHandler Me._Window.Closing, AddressOf Me.CaptureValues
            Me._Window.ShowDialog()
            Return Me.ButtonPress
        End Function

        Private Sub CaptureValues(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
            Me._Button = Me._Window.ButtonValue
        End Sub

    End Class

    Public Class ErrorBox

        Dim _Message As String
        Dim _Title As String
        Dim _Details As String
        Dim _Severity As ErrorType
        Dim _Button1 As String
        Dim _Button2 As String
        Dim _Window As New ErrorPop
        Dim _Button As Response
        Dim _EmailAdmin As Boolean = False
        Dim _LogError As Boolean = False

        Public Enum ErrorType
            Confirmation = 0
            Minor = 1
            Critical = 2
        End Enum

        Public ReadOnly Property ButtonPress() As Response
            Get
                Return _Button
            End Get
        End Property

        Public ReadOnly Property NotifyAdmin() As Boolean
            Get
                Return Me._EmailAdmin
            End Get
        End Property

        Public ReadOnly Property LogError() As Boolean
            Get
                Return Me._LogError
            End Get
        End Property

        Public Sub New(ByVal Message As String, Optional ByVal Title As String = "Gravity Error", _
        Optional ByVal Details As String = "", Optional ByVal Severity As Integer = 1, _
        Optional ByVal Button1Text As String = "&Continue", Optional ByVal Button2Text As String = "&Close Program")
            Me.SetValues(Message, Title, Details, Severity, Button1Text, Button2Text)
            Me.ShowDialog()
        End Sub

        Public Sub New()
            ' nothing
        End Sub

        Private Sub SetValues(ByVal Message As String, ByVal Title As String, ByVal Details As String, _
        ByVal Severity As ErrorType, ByVal Button1Text As String, ByVal Button2Text As String)
            Me._Message = Message
            Me._Title = Title
            Me._Details = Details
            Me._Severity = Severity
            Me._Button1 = Button1Text
            Me._Button2 = Button2Text
        End Sub

        Public Sub ShowDialog(ByVal Message As String, Optional ByVal Title As String = "Gravity Error", _
        Optional ByVal Details As String = "", Optional ByVal Severity As ErrorType = 1, _
        Optional ByVal Button1Text As String = "&Continue", Optional ByVal Button2Text As String = "&Close Program")
            Me.SetValues(Message, Title, Details, Severity, Button1Text, Button2Text)
            Me.ShowDialog()
        End Sub

        Public Sub ShowDialog()
            Me._Window.lblCaption.Text = Me._Message
            Me._Window.Text = Me._Title
            Me._Window.txtDetails.Text = Me._Details
            Me._Window.ubtnContinue.Text = Me._Button1
            Me._Window.ubtnCancel.Text = Me._Button2
            If Me._Details.Length = 0 Then
                Me._Window.txtDetails.Visible = False
                Me._Window.lblCaption.Height += 30
                'Me._Window.btnContinue.Top -= 20
                'Me._Window.btnClose.Top -= 20
                'Me._Window.Height -= 20
            End If
            If Me._Severity = ErrorType.Minor Then
                Me._Window.ubtnCancel.Visible = False
                Me._Window.uchkLog.Checked = False
                Me._Window.uchkNotify.Checked = False
                Me._Window.uchkLog.Visible = False
                Me._Window.uchkNotify.Visible = False
            ElseIf Me._Severity = ErrorType.Critical Then
                Me._Window.uchkLog.Checked = True
                Me._Window.uchkNotify.Checked = True
                Me._Window.uchkLog.Visible = True
                Me._Window.uchkNotify.Visible = True
            ElseIf Me._Severity = ErrorType.Confirmation Then
                Me._Window.uchkLog.Checked = False
                Me._Window.uchkNotify.Checked = False
                Me._Window.uchkLog.Visible = False
                Me._Window.uchkNotify.Visible = False
            End If
            AddHandler Me._Window.Closing, AddressOf Me.CaptureValues
            Me._Window.ShowDialog()
        End Sub

        Private Sub CaptureValues(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
            Me._EmailAdmin = Me._Window.uchkNotify.Checked
            Me._LogError = Me._Window.uchkLog.Checked
            Me._Button = Me._Window.ButtonValue
        End Sub

    End Class

    Public Class AskBox

        Dim _Message As String
        Dim _Title As String
        Dim _Cancel As Boolean
        Dim _Button1 As String
        Dim _Button2 As String
        Dim _Button3 As String
        Dim _Window As New AskWindow
        Dim _Button As Response

        Public ReadOnly Property ButtonPress() As Response
            Get
                Return _Button
            End Get
        End Property

        Public Sub New()

        End Sub

        Public Sub New(ByVal Message As String, Optional ByVal Title As String = "Gravity Question", _
        Optional ByVal Cancel As Integer = 1, Optional ByVal Button1Text As String = "&No", _
        Optional ByVal Button2Text As String = "&Yes", Optional ByVal Button3Text As String = "&Cancel")
            Me.SetValues(Message, Title, Cancel, Button1Text, Button2Text, Button3Text)
            Me.ShowDialog()
        End Sub

        Private Sub SetValues(ByVal Message As String, ByVal Title As String, ByVal Cancel As Integer, _
        ByVal Button1Text As String, ByVal Button2Text As String, ByVal Button3Text As String)
            Me._Message = Message
            Me._Title = Title
            Me._Cancel = Cancel
            Me._Button1 = Button1Text
            Me._Button2 = Button2Text
            Me._Button3 = Button3Text
        End Sub

        Public Sub ShowDialog(ByVal Message As String, Optional ByVal Title As String = "Gravity Question", _
        Optional ByVal Cancel As Integer = 1, Optional ByVal Button1Text As String = "&No", _
        Optional ByVal Button2Text As String = "&Yes", Optional ByVal Button3Text As String = "&Cancel")
            Me.SetValues(Message, Title, Cancel, Button1Text, Button2Text, Button3Text)
            Me.ShowDialog()
        End Sub

        Public Sub ShowDialog()
            Me._Window.lblCaption.Text = Me._Message
            Me._Window.Text = Me._Title
            Me._Window.ubtnNo.Text = Me._Button1
            Me._Window.ubtnYes.Text = Me._Button2
            Me._Window.ubtnCancel.Text = Me._Button3
            If Not Me._Cancel Then
                Me._Window.ubtnCancel.Visible = False
            End If
            AddHandler Me._Window.Closing, AddressOf Me.CaptureValues
            Me._Window.ShowDialog()
        End Sub

        Private Sub CaptureValues(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
            Me._Button = Me._Window.ButtonValue
        End Sub

    End Class

    Public Class InputBox

        Dim _Message As String
        Dim _Title As String
        Dim _DefaultText As String
        Dim _Button1 As String
        Dim _Button2 As String
        Dim _Window As New GravityInput
        Dim _Button As Response
        Dim _Text As String
        Dim _Mask As String = Nothing

        Public ReadOnly Property ButtonPress() As Response
            Get
                Return _Button
            End Get
        End Property

        Public ReadOnly Property Text() As String
            Get
                Return _Text
            End Get
        End Property

        Public Sub New(ByVal Message As String, Optional ByVal Title As String = "Gravity Input", _
        Optional ByVal DefaultText As String = "", _
        Optional ByVal Button1Text As String = "&OK", Optional ByVal Button2Text As String = "&Cancel")
            Me.SetValues(Message, Title, DefaultText, Button1Text, Button2Text)
            Me.ShowDialog()
        End Sub

        Public Sub New()

        End Sub

        Public Sub SetValues(ByVal Message As String, ByVal Title As String, Optional ByVal DefaultText As String = "", _
        Optional ByVal Button1Text As String = "", Optional ByVal Button2Text As String = "")
            Me._Message = Message
            Me._Title = Title
            Me._DefaultText = DefaultText
            Me._Button1 = Button1Text
            Me._Button2 = Button2Text
        End Sub

        Public Sub SetMask(ByVal Mask As String)
            Me._Mask = Mask
        End Sub

        Public Function ShowDialog(ByVal Message As String, Optional ByVal Title As String = "Gravity Input", _
        Optional ByVal DefaultText As String = "", _
        Optional ByVal Button1Text As String = "&OK", Optional ByVal Button2Text As String = "&Cancel") As Response
            Me.SetValues(Message, Title, DefaultText, Button1Text, Button2Text)
            Me.ShowDialog()
            Return Me.ButtonPress
        End Function

        Public Function ShowDialog() As Response
            Me._Window.lblCaption.Text = Me._Message
            Me._Window.Text = Me._Title
            If Me._Mask = Nothing Then
                Me._Window.txtInput.Text = Me._DefaultText
                Me._Window.txtMaskedInput.Visible = False
            Else
                Me._Window.txtMaskedInput.Text = Me._DefaultText
                'Me._Window.txtMaskedInput.Mask = Me._Mask
                Me._Window.txtInput.Visible = False
            End If
            Me._Window.ubtnOK.Text = Me._Button1
            Me._Window.ubtnCancel.Text = Me._Button2
            AddHandler Me._Window.Closing, AddressOf Me.CaptureValues
            Me._Window.ShowDialog()
            Return Me.ButtonPress
        End Function

        Private Sub CaptureValues(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
            Me._Button = Me._Window.ButtonValue
            If Me._Mask = Nothing Then
                Me._Text = Me._Window.txtInput.Text
            Else
                Me._Text = Me._Window.txtMaskedInput.Text
            End If
        End Sub

    End Class


    Public Class EmailInputBox

        Dim _Title As String
        Dim _Window As New frmEmailForm
        Dim _Button As Response
        Dim _Mask As String = Nothing

        Public Email As String = ""
        Public Subject As String = ""
        Public Name As String = ""

        Public ReadOnly Property ButtonPress() As Response
            Get
                Return _Button
            End Get
        End Property

        Public Sub New()

        End Sub

        Public Sub ShowDialog()
            Me._Window.txtEmail.Text = Me.Email
            Me._Window.txtName.Text = Me.Name
            Me._Window.txtSubject.Text = Me.Subject
            Me._Window.Text = Me._Title
            AddHandler Me._Window.Closing, AddressOf Me.CaptureValues
            Me._Window.ShowDialog()
        End Sub

        Private Sub CaptureValues(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
            Me._Button = Me._Window.ButtonValue
            Me.Email = Me._Window.txtEmail.Text
            Me.Subject = Me._Window.txtSubject.Text
            Me.Name = Me._Window.txtName.Text
        End Sub

    End Class


    Public Class FormattedEditorBox

        Dim _Title As String = "Text Formatter"
        Dim _DefaultText As String = ""
        Dim _Button1 As String = "OK"
        Dim _Button2 As String = "Cancel"
        Dim _Window As New FormattedTextEditor
        Dim _Button As Response
        Dim _Text As String

        Public IsReadOnly As Boolean = False

        Public Event OKClicked(ByVal win As FormattedEditorBox)
        Public Event CancelClicked(ByVal win As FormattedEditorBox)

        Public ReadOnly Property ButtonPress() As Response
            Get
                Return _Button
            End Get
        End Property

        Public ReadOnly Property Text() As String
            Get
                Return _Text
            End Get
        End Property

        Public Sub New(ByVal Title As String, Optional ByVal DefaultText As String = "", Optional ByVal Button1Text As String = "&OK", Optional ByVal Button2Text As String = "&Cancel")
            Me.SetValues(Title, DefaultText, Button1Text, Button2Text)
            Me.ShowDialog()
        End Sub

        Public Sub New()

        End Sub

        Public Sub SetValues(ByVal Title As String, ByVal DefaultText As String, ByVal Button1Text As String, ByVal Button2Text As String)
            Me._Title = Title
            Me._DefaultText = DefaultText
            Me._Button1 = Button1Text
            Me._Button2 = Button2Text
        End Sub

        Public Sub ShowDialog(ByVal Title As String, Optional ByVal DefaultText As String = "", Optional ByVal Button1Text As String = "&OK", Optional ByVal Button2Text As String = "&Cancel")
            Me.SetValues(Title, DefaultText, Button1Text, Button2Text)
            Me.ShowDialog()
        End Sub

        Public Sub ShowDialog()
            Me._Window.Text = Me._Title
            Me._Window.uftxtBody.Value = Me._DefaultText
            Me._Window.btnOK.Text = Me._Button1
            Me._Window.btnCancel.Text = Me._Button2
            Me._Window.uftxtBody.ReadOnly = Me.IsReadOnly
            If Me.IsReadOnly Then
                Me._Window.ToolStrip2.Visible = False
            End If
            AddHandler Me._Window.Closing, AddressOf Me.CaptureValues
            AddHandler Me._Window.OKPressed, AddressOf Me.ClickOK
            AddHandler Me._Window.CancelPressed, AddressOf Me.ClickCancel
            Me._Window.ShowDialog()
        End Sub

        Private Sub CaptureValues(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
            Me._Button = Me._Window.ButtonValue
            Me._Text = Me._Window.uftxtBody.Value
        End Sub

        Private Sub ClickOK(ByVal win As FormattedTextEditor)
            RaiseEvent OKClicked(Me)
        End Sub

        Private Sub ClickCancel(ByVal win As FormattedTextEditor)
            RaiseEvent CancelClicked(Me)
        End Sub

    End Class


    Public Class SelectBox

        Dim _Message As String
        Dim _Title As String
        Dim _Button1 As String
        Dim _Button2 As String
        Dim _Window As New frmSelect
        Dim _Button As Response
        Dim _Value As String = ""

        Public ReadOnly Property ButtonPress() As Response
            Get
                Return _Button
            End Get
        End Property

        Public ReadOnly Property Value() As String
            Get
                Return _Value
            End Get
        End Property

        Public Property DropDown() As Infragistics.Win.UltraWinGrid.UltraCombo
            Get
                Return Me._Window.ucboChoice
            End Get
            Set(ByVal Combo As Infragistics.Win.UltraWinGrid.UltraCombo)
                Me._Window.ucboChoice = Combo
            End Set
        End Property

        Public Sub New(ByVal Message As String, Optional ByVal Title As String = "Gravity Input", _
        Optional ByVal Button1Text As String = "&OK", Optional ByVal Button2Text As String = "&Cancel")
            Me.SetValues(Message, Title, Button1Text, Button2Text)
        End Sub

        Public Sub New()

        End Sub

        Public Sub SetValues(ByVal Message As String, ByVal Title As String, _
        ByVal Button1Text As String, ByVal Button2Text As String)
            Me._Message = Message
            Me._Title = Title
            Me._Button1 = Button1Text
            Me._Button2 = Button2Text
        End Sub

        Public Sub ShowDialog(ByVal Message As String, Optional ByVal Title As String = "Gravity Input", _
        Optional ByVal Button1Text As String = "&OK", Optional ByVal Button2Text As String = "&Cancel")
            Me.SetValues(Message, Title, Button1Text, Button2Text)
            Me.ShowDialog()
        End Sub

        Public Sub ShowDialog()
            Me._Window.lblCaption.Text = Me._Message
            Me._Window.Text = Me._Title
            Me._Window.ubtnOK.Text = Me._Button1
            Me._Window.ubtnCancel.Text = Me._Button2
            AddHandler Me._Window.Closing, AddressOf Me.CaptureValues
            Me._Window.ShowDialog()
        End Sub

        Private Sub CaptureValues(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
            Me._Button = Me._Window.ButtonValue
            Me._Value = Me._Window.ucboChoice.Value
        End Sub

    End Class

    Public Class DateRange

        Dim _Message As String
        Dim _Title As String
        Dim _Start As Date = Today
        Dim _End As Date = Today
        Dim _Button1 As String
        Dim _Button2 As String
        Dim _Window As New frmDateBox
        Dim _Button As Response

        Public IsDateRange As Boolean = True

        Public ReadOnly Property ButtonPress() As Response
            Get
                Return _Button
            End Get
        End Property

        Public Property StartDate() As Date
            Get
                Return Me._Start
            End Get
            Set(ByVal Value As Date)
                Me._Start = Value
            End Set
        End Property

        Public Property EndDate() As Date
            Get
                Return Me._End
            End Get
            Set(ByVal Value As Date)
                Me._End = Value
            End Set
        End Property

        Public Sub New(ByVal Message As String, Optional ByVal Title As String = "Gravity Date Range", _
        Optional ByVal Button1Text As String = "&OK", Optional ByVal Button2Text As String = "&Cancel")
            Me.SetValues(Message, Title, Button1Text, Button2Text)
            Me.ShowDialog()
        End Sub

        Public Sub New()

        End Sub

        Public Sub SetValues(ByVal Message As String, ByVal Title As String, _
        ByVal Button1Text As String, ByVal Button2Text As String)
            Me._Message = Message
            Me._Title = Title
            Me._Button1 = Button1Text
            Me._Button2 = Button2Text
        End Sub

        Public Sub ShowDialog(ByVal Message As String, Optional ByVal Title As String = "Gravity Date Range", _
        Optional ByVal Button1Text As String = "&OK", Optional ByVal Button2Text As String = "&Cancel")
            Me.SetValues(Message, Title, Button1Text, Button2Text)
            Me.ShowDialog()
        End Sub

        Public Sub ShowDialog()
            Me._Window.lblCaption.Text = Me._Message
            Me._Window.Text = Me._Title
            Me._Window.udtpStart.Value = Me._Start
            Me._Window.udtpEnd.Value = Me._End
            Me._Window.ubtnOK.Text = Me._Button1
            Me._Window.ubtnCancel.Text = Me._Button2
            If Not Me.IsDateRange Then
                Me._Window.pnlBetween.Visible = False
                Me._Window.udtpEnd.Visible = False
            End If
            AddHandler Me._Window.Closing, AddressOf Me.CaptureValues
            Me._Window.ShowDialog()
        End Sub

        Private Sub CaptureValues(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
            Me._Button = Me._Window.ButtonValue
            Me._Start = Me._Window.udtpStart.Value
            Me._End = Me._Window.udtpEnd.Value
        End Sub

    End Class

    Public Class MultiInputBox

        Dim _Message As String
        Dim _Title As String
        Dim _DefaultText As String
        Dim _Button1 As String
        Dim _Button2 As String
        Dim _Label1 As String
        Dim _Label2 As String
        Dim _Label3 As String
        Dim _Default1 As String
        Dim _Default2 As String
        Dim _Default3 As String
        Dim _Window As New MultiInput
        Dim _Button As Response
        Dim _Text1 As String
        Dim _Text2 As String
        Dim _Text3 As String
        Dim _Mask1 As Boolean
        Dim _Mask2 As Boolean
        Dim _Mask3 As Boolean

        Public ReadOnly Property ButtonPress() As Response
            Get
                Return _Button
            End Get
        End Property

        Public ReadOnly Property Text() As String()
            Get
                Dim ReturnVal(2) As String
                ReturnVal(0) = _Text1
                ReturnVal(1) = _Text2
                ReturnVal(2) = _Text3
                Return ReturnVal
            End Get
        End Property

        Public Sub New(ByVal Message As String, Optional ByVal Title As String = "Gravity Multi-Input", _
        Optional ByVal Label1 As String = "", Optional ByVal Label2 As String = "", Optional ByVal Label3 As String = "", _
        Optional ByVal DefaultText1 As String = "", Optional ByVal DefaultText2 As String = "", Optional ByVal DefaultText3 As String = "", _
        Optional ByVal Button1Text As String = "&OK", Optional ByVal Button2Text As String = "&Cancel")
            Me.SetValues(Message, Title, Label1, DefaultText1, Label2, DefaultText2, Label3, DefaultText3, Button1Text, Button2Text)
            Me.ShowDialog()
        End Sub

        Private Sub SetValues(ByVal Message As String, ByVal Title As String, _
        ByVal Label1 As String, ByVal DefaultText1 As String, _
        ByVal Label2 As String, ByVal DefaultText2 As String, _
        ByVal Label3 As String, ByVal DefaultText3 As String, _
        ByVal Button1Text As String, ByVal Button2Text As String)
            Me._Message = Message
            Me._Title = Title
            Me._Default1 = DefaultText1
            Me._Default2 = DefaultText2
            Me._Default3 = DefaultText3
            Me._Label1 = Label1
            Me._Label2 = Label2
            Me._Label3 = Label3
            Me._Button1 = Button1Text
            Me._Button2 = Button2Text
        End Sub

        Public Sub SetPasswordMask(ByVal Field As Integer)
            Select Case Field
                Case 1
                    Me._Mask1 = True
                Case 2
                    Me._Mask2 = True
                Case 3
                    Me._Mask3 = True
            End Select
        End Sub

        Public Sub ShowDialog(ByVal Message As String, Optional ByVal Title As String = "Gravity Input", _
        Optional ByVal Label1 As String = "", Optional ByVal Label2 As String = "", Optional ByVal Label3 As String = "", _
        Optional ByVal DefaultText1 As String = "", Optional ByVal DefaultText2 As String = "", Optional ByVal DefaultText3 As String = "", _
        Optional ByVal Button1Text As String = "&OK", Optional ByVal Button2Text As String = "&Cancel")
            Me.SetValues(Message, Title, Label1, DefaultText1, Label2, DefaultText2, Label3, DefaultText3, Button1Text, Button2Text)
            Me.ShowDialog()
        End Sub

        Public Sub ShowDialog()
            Me._Window.lblCaption.Text = Me._Message
            Me._Window.Text = Me._Title
            Me._Window.Label1.Text = Me._Label1
            Me._Window.TextBox1.Text = Me._Text1
            Me._Window.Label2.Text = Me._Label2
            Me._Window.TextBox2.Text = Me._Text2
            If Me._Label3.Length = 0 Then
                Me._Window.Label3.Visible = False
                Me._Window.TextBox3.Visible = False
            Else
                Me._Window.Label3.Text = Me._Label3
                Me._Window.TextBox3.Text = Me._Text3
            End If
            ' Set password masks
            If Me._Mask1 Then
                Me._Window.TextBox1.PasswordChar = "*"
            End If
            If Me._Mask2 Then
                Me._Window.TextBox2.PasswordChar = "*"
            End If
            If Me._Mask3 Then
                Me._Window.TextBox3.PasswordChar = "*"
            End If
            ' Button text
            Me._Window.ubtnOK.Text = Me._Button1
            Me._Window.ubtnCancel.Text = Me._Button2
            ' Set capture values on window close
            AddHandler Me._Window.Closing, AddressOf Me.CaptureValues
            ' Show window
            Me._Window.ShowDialog()
        End Sub

        Private Sub CaptureValues(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
            Me._Button = Me._Window.ButtonValue
            Me._Text1 = Me._Window.TextBox1.Text
            Me._Text2 = Me._Window.TextBox2.Text
            Me._Text3 = Me._Window.TextBox3.Text
        End Sub

    End Class

    Public Class ChooseTemplateDialog

        Public Title As String = ""
        Public Database As MyCore.Data.EasySql
        Public WindowMode As Mode = Mode.OpenTemplate
        Public ButtonPress As Response = Response.Cancel

        Dim WithEvents _Window As New Browse

        Public SelectedFolderID As Integer = 0
        Public SelectedTemplateID As Integer = 0
        Public SelectedTemplateName As String = ""
        Public SelectedTemplateHtml As String = ""

        Public Enum Mode
            SaveTemplate = 0
            OpenTemplate = 1
        End Enum

        Public Sub New(ByRef db As MyCore.Data.EasySql, ByVal m As Mode)
            Me.Database = db
            Me.WindowMode = m
            Me._Window.Controller = Me
        End Sub

        Public Function ShowDialog() As Response
            If Me.WindowMode = Mode.SaveTemplate Then
                Me._Window.pnlNew.Visible = True
                If Me.Title.Length = 0 Then
                    Me._Window.Text = "Save Template"
                Else
                    Me._Window.Text = Me.Title
                End If
            Else
                Me._Window.pnlNew.Visible = False
                If Me.Title.Length = 0 Then
                    Me._Window.Text = "Choose Template"
                Else
                    Me._Window.Text = Me.Title
                End If
            End If
            Me._Window.InitialFolder = Me.SelectedFolderID
            Me._Window.txtName.Text = Me.SelectedTemplateName
            Me._Window.ShowDialog()
            Return Me.ButtonPress
        End Function

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class

End Namespace