Public Class frmLogin

    Public ParentWin As MyCore.Plugins.IHost

    Dim Crypt As New MyCore.Utility.SimpleEncryption(1)
    Dim Strikes As Integer = 0

    Public Event Canceled(ByVal sender As frmLogin)
    Public Event LoggedIn(ByVal sender As frmLogin, ByVal UserName As String)
    Public Event LoginRejected(ByVal sender As frmLogin)
    Public Event AttemptingLogin(ByVal sender As frmLogin, ByVal UserName As String, ByVal PasswordHash As String)

    Public DoAuthentication As Boolean = True

    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.utxtUser.Focus()
        If Me.DoAuthentication Then
            Dim AllowAuto As Boolean = False
            Try
                AllowAuto = IIf(Me.ParentWin.Config.GetSetting("Allow Auto Login", "Y") = "Y", True, False)
            Catch ex As Exception
                ' Nothing
            End Try
            If AllowAuto Then
                If Me.ParentWin.Config.File = Me.ParentWin.SettingsLocal.GetSetting("last_crm_file") Then
                    If Me.ParentWin.SettingsLocal.GetSetting("remember_login", "N").ToUpper = "Y" Then
                        Me.utxtUser.Text = Me.ParentWin.SettingsLocal.GetSetting("last_username")
                        Me.utxtPassword.Text = Me.Crypt.Decrypt(Me.ParentWin.SettingsLocal.GetSetting("last_password"))
                    End If
                End If
            Else
                Me.uchkRemember.Enabled = False
            End If
        Else
            Me.uchkRemember.Visible = False
        End If
    End Sub

    Private Sub AttemptLogin()
        Dim PassHash As String = MyCore.Utility.Hash.MD5(Me.utxtPassword.Text)
        Dim Sql As String = "SELECT COUNT(id) FROM employee WHERE"
        Sql &= " windows_user=" & Me.ParentWin.Database.Escape(Me.utxtUser.Text)
        If PassHash.Length > 12 Then
            Sql &= " AND password_hash LIKE " & Me.ParentWin.Database.Escape(PassHash & "%")
        Else
            Sql &= " AND password_hash=" & Me.ParentWin.Database.Escape(PassHash)
        End If
        Try
            If Me.ParentWin.Database.GetOne(Sql) > 0 Then
                Me.Hide()
                RaiseEvent LoggedIn(Me, Me.utxtUser.Text)
                If Me.uchkRemember.Checked Then
                    Me.ParentWin.SettingsLocal.SaveSetting("remember_login", "Y")
                    Me.ParentWin.SettingsLocal.SaveSetting("last_username", Me.utxtUser.Text)
                    Me.ParentWin.SettingsLocal.SaveSetting("last_password", Me.Crypt.Encrypt(Me.utxtPassword.Text))
                End If
                Me.Close()
            Else
                Dim ErrorWin As New MyCore.Gravity.ErrorBox("User name or password was incorrect.", "Login Error")
                If Me.Strikes < 3 Then
                    Me.Strikes += 1
                Else
                    RaiseEvent LoginRejected(Me)
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            If ex.ToString.Contains("An error has occurred while establishing a connection to the server.") Then
                Dim ErrorWin As New MyCore.Gravity.ErrorBox("Gravity could not connect to your server.  Make sure the server is up and that the database and connection are configured correctly.", "Login Error")
            ElseIf ex.ToString.Contains("timeout") Then
                Dim ErrorWin As New MyCore.Gravity.ErrorBox("Gravity encountered a timeout error when trying to get login information from the database. There may be something running on your server that is hogging the system resources. Try again later or check your server.", "Login Error")
            Else
                Dim ErrorWin As New MyCore.Gravity.ErrorBox("Login failed due to unknown error.", "Login Error", ex.ToString)
            End If
        End Try
    End Sub

    Private Sub lblLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblLogin.Click
        RaiseEvent AttemptingLogin(Me, Me.utxtUser.Text, MyCore.Utility.Hash.MD5(Me.utxtPassword.Text))
        If Me.DoAuthentication Then
            Me.AttemptLogin()
        End If
    End Sub

    Private Sub Cancel()
        RaiseEvent Canceled(Me)
        Me.Close()
    End Sub

    Private Sub utxtPassword_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles utxtPassword.KeyUp, utxtUser.KeyUp
        If e.KeyCode = Keys.Enter Then
            Me.lblLogin_Click(Me, New EventArgs)
        ElseIf e.KeyCode = Keys.Escape Then
            Me.Cancel()
        End If
    End Sub

    Private Sub lnkContinue_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkContinue.LinkClicked
        Me.Cancel()
    End Sub

End Class