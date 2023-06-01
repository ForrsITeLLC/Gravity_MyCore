Public Class frmEmailForm

    Public ButtonValue As Gravity.Response = Gravity.Response.OK

    Private Sub GravityInput_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txtEmail.Focus()
    End Sub

    Private Sub ubtnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnOK.Click
        Me.ButtonValue = Gravity.Response.OK
        Me.Close()
    End Sub

    Private Sub ubtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnCancel.Click
        Me.ButtonValue = Gravity.Response.Cancel
        Me.Close()
    End Sub

End Class