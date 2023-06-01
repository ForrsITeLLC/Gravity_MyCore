Public Class FormattedTextEditor

    Public ButtonValue As Gravity.Response = Gravity.Response.OK

    Public Event OKPressed(ByVal win As FormattedTextEditor)
    Public Event CancelPressed(ByVal win As FormattedTextEditor)

    Private Sub GravityInput_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.uftxtBody.Focus()
    End Sub

    Private Sub ubtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Me.ButtonValue = Gravity.Response.OK
        RaiseEvent OKPressed(Me)
        Me.Close()
    End Sub

    Private Sub ubtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.ButtonValue = Gravity.Response.Cancel
        RaiseEvent CancelPressed(Me)
        Me.Close()
    End Sub

    Private Sub tbtnLink_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbtnLink.Click
        If Me.uftxtBody.EditInfo.CanMakeSelectionIntoLink Then
            Me.uftxtBody.EditInfo.ShowLinkDialog()
        Else
            MessageBox.Show("Selected text cannot be made a link.  Either no text is selected or areas that are already linked and not linked overlap. Select a valid text area and try again.")
        End If
    End Sub

    Private Sub tbtnBold_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbtnBold.Click
        Me.uftxtBody.EditInfo.ShowFontDialog()
    End Sub

    Private Sub btnImg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImg.Click
        Me.uftxtBody.EditInfo.ShowImageDialog()
    End Sub

End Class