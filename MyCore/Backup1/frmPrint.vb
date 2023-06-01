Public Class frmPrint

    Public File As String
    Public Action As ActionType
    Public AskIfSuccessful As Boolean = True

    Public Event PrintSuccessful()
    Public Event PrintFailed()

    Public Enum ActionType
        SilentPrint = 1
        PrintPreview = 2
        PageSettings = 3
        PrintDialog = 4
    End Enum

    Public Sub New(ByVal FilePath As String, Optional ByVal a As ActionType = ActionType.PrintDialog, Optional ByVal AskSuccessFail As Boolean = True)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.File = FilePath
        Me.Action = a
        Me.AskIfSuccessful = AskSuccessFail

    End Sub

    Private Sub Print_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Me.Action = ActionType.PageSettings Then
            Me.Label1.Text = "Changing page settings"
        ElseIf Me.Action = ActionType.PrintPreview Then
            Me.Label1.Text = "Generating print preview"
        End If
        Timer1.Start()
    End Sub

    Private Sub Print() Handles Timer1.Tick
        Timer1.Stop()
        ' Open File
        Me.Label1.Text = "Loading document..."
        Me.WebBrowser1.Navigate(Me.File)
    End Sub

    Private Sub ubtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnSuccess.Click
        RaiseEvent PrintSuccessful()
        Me.Close()
    End Sub

    Private Sub ubtnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnPrint.Click
        Me.WebBrowser1.ShowPrintDialog()
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        If Me.WebBrowser1.Document.Url.IsFile Then
            ' Print
            Me.Label1.Text = "Sending to printer..."
            If Me.Action = ActionType.PageSettings Then
                Me.WebBrowser1.ShowPageSetupDialog()
                Me.Label1.Text = "Did you print the document?"
                Me.ubtnPrint.Text = "Print"
            ElseIf Me.Action = ActionType.PrintPreview Then
                Me.WebBrowser1.ShowPrintPreviewDialog()
                Me.Label1.Text = "Did you print the document?"
                Me.ubtnPrint.Text = "Print"
            ElseIf Me.Action = ActionType.PrintDialog Then
                Me.WebBrowser1.ShowPrintDialog()
                Me.Label1.Text = "Was the print successful?"
            Else
                Me.Label1.Text = "Printing..."
                Me.WebBrowser1.Print()
                Me.Close()
            End If
            If Not Me.AskIfSuccessful Then
                Me.Label1.Text = "Done printing?"
                Me.ubtnFailed.Visible = False
                Me.ubtnSuccess.Text = "Finished"
            End If
            Me.pnlOptions.Visible = True
        End If
    End Sub

    Private Sub ubtnFailed_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ubtnFailed.Click
        RaiseEvent PrintFailed()
        Me.Close()
    End Sub

End Class