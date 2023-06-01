Imports System.Net.Mail

Namespace Email

    Public Class MailSettings

        Public UserName As String = ""
        Public Password As String = ""
        Public Domain As String = ""
        Public Host As String = ""
        Public Port As Integer = 23
        Public EnableSSL As Boolean = False
        Public RequireAuthentication As Boolean = True

    End Class

    Public Class Message

        Dim Mail As New MailMessage
        Dim Settings As MailSettings

        Public FromName As String = ""
        Public ReplyToName As String = ""

        Public FromAddress As String = ""
        Public ReplyToAddress As String = ""

        Public Subject As String = ""
        Public Body As String = ""
        Public HtmlBody As String = ""

        Public Sub New(ByVal Settings As MailSettings)
            Me.Settings = Settings
        End Sub

        Public Sub AddToAddress(ByVal Email As String, ByVal Name As String)
            Me.Mail.To.Add(New MailAddress(Email, Name))
        End Sub

        Public Sub AddCCAddress(ByVal Email As String, ByVal Name As String)
            Me.Mail.CC.Add(New MailAddress(Email, Name))
        End Sub

        Public Sub AddBCCAddress(ByVal Email As String, ByVal Name As String)
            Me.Mail.Bcc.Add(New MailAddress(Email, Name))
        End Sub

        Public Sub AddAttachment(ByVal FileName As String)
            Me.Mail.Attachments.Add(New Attachment(FileName))
        End Sub

        Public Sub ClearAttachments()
            Me.Mail.Attachments.Clear()
        End Sub

        Public Sub ClearToAddresses()
            Me.Mail.To.Clear()
        End Sub

        Public Sub Send()

            Dim Smtp As New SmtpClient

            ' Configure smtp options
            Smtp.Host = Me.Settings.Host
            Smtp.Port = Me.Settings.Port
            Smtp.EnableSsl = Me.Settings.EnableSSL
            Smtp.Timeout = 100000
            If Me.Settings.RequireAuthentication Then
                Dim Cred As New System.Net.NetworkCredential(Me.Settings.UserName, Me.Settings.Password, Me.Settings.Domain)
                Smtp.Credentials = Cred
            End If


            ' Configure message options
            Mail.From = New MailAddress(Me.FromAddress, Me.FromName)
            If Me.ReplyToAddress.Length > 0 Then
                Mail.ReplyTo = New MailAddress(Me.ReplyToAddress, Me.ReplyToName)
            End If
            Mail.Subject = Me.Subject
            Mail.Body = Me.Body

            ' Send email
            Smtp.Send(Mail)

        End Sub


    End Class




End Namespace