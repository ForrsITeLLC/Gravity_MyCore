Imports System.Windows.Forms

Public Class frmNewUpdate

    Public Url As String = ""
    Public Event ExitProgram()
    Public Event RemindMeLater()

    Private Sub btnLater_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLater.Click
        RaiseEvent RemindMeLater()
        Me.Hide()
    End Sub

    Private Sub btnNow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNow.Click

        ' Download file
        Dim Dest As String = My.Application.Info.DirectoryPath & "\update.exe"
        Try
            My.Computer.Network.DownloadFile(Url, Dest, "", "", True, 100000, True)
        Catch ex As Exception
            Dim Err As New MyCore.Gravity.ErrorBox
            Err.ShowDialog("Error downloading file.  Try again later.", "Download Error", ex.ToString)
            Me.Hide()
            Exit Sub
        End Try

        ' Go ahead?
        If MessageBox.Show("Please confirm that you are ready to install the update.  Be sure you have saved all work.  Proceed?", "Install Confirmation", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
            Me.InstallNow()
        Else
            Me.Cancel()
        End If
    End Sub

    Private Sub InstallNow()

        ' Unsynced items
        Dim UnsyncedPath As String = My.Computer.FileSystem.SpecialDirectories.ProgramFiles & "\EVware\Gravity CRM\Service\unsynced.txt"
        If My.Computer.FileSystem.FileExists(UnsyncedPath) Then
            If System.Windows.Forms.MessageBox.Show("UNSYNCED ITEMS FOUND!!!  Are you sure you want to continue?", "Unsynced Items", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
                Me.Cancel()
            End If
        End If

        ' Downloaded file
        Dim Dest As String = My.Application.Info.DirectoryPath & "\update.exe"

        ' Proceed with installation
        Me.Hide()
        System.Diagnostics.Process.Start(Dest)
        RaiseEvent ExitProgram()

    End Sub

    Private Sub Cancel()
        Me.Hide()
    End Sub

End Class