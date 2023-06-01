Public Class ArrowButtonNext

    Public Event ButtonClick(ByVal sender As Object, ByVal e As EventArgs)

    Public Property ButtonText() As String
        Get
            Return Me.Button.Text
        End Get
        Set(ByVal value As String)
            Me.Button.Text = value
        End Set
    End Property

    Public Property ButtonColor1() As Color
        Get
            Return Me.Button.Appearance.BackColor
        End Get
        Set(ByVal value As Color)
            Me.Button.Appearance.BackColor = value
        End Set
    End Property

    Public Property ButtonColor2() As Color
        Get
            Return Me.Button.Appearance.BackColor2
        End Get
        Set(ByVal value As Color)
            Me.Button.Appearance.BackColor2 = value
        End Set
    End Property

    Private Sub Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button.Click
        RaiseEvent ButtonClick(Me, New System.EventArgs)
    End Sub

End Class
