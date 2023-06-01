Public Class SignatureBox
    Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'UserControl overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents lnkClearSignature As System.Windows.Forms.LinkLabel
    Friend WithEvents InputArea As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lnkClearSignature = New System.Windows.Forms.LinkLabel
        Me.InputArea = New System.Windows.Forms.PictureBox
        Me.SuspendLayout()
        '
        'lnkClearSignature
        '
        Me.lnkClearSignature.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lnkClearSignature.Location = New System.Drawing.Point(0, 112)
        Me.lnkClearSignature.Name = "lnkClearSignature"
        Me.lnkClearSignature.Size = New System.Drawing.Size(344, 16)
        Me.lnkClearSignature.TabIndex = 51
        Me.lnkClearSignature.TabStop = True
        Me.lnkClearSignature.Text = "Clear Signature"
        Me.lnkClearSignature.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'InputArea
        '
        Me.InputArea.BackColor = System.Drawing.Color.White
        Me.InputArea.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.InputArea.Dock = System.Windows.Forms.DockStyle.Fill
        Me.InputArea.Location = New System.Drawing.Point(0, 0)
        Me.InputArea.Name = "InputArea"
        Me.InputArea.Size = New System.Drawing.Size(344, 112)
        Me.InputArea.TabIndex = 52
        Me.InputArea.TabStop = False
        '
        'SignatureBox
        '
        Me.Controls.Add(Me.InputArea)
        Me.Controls.Add(Me.lnkClearSignature)
        Me.Name = "SignatureBox"
        Me.Size = New System.Drawing.Size(344, 128)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Event InputCleared()
    Public Event Signing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
    Public Event StartSigning(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
    Public Event StopSigning(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
    Public Event SignedStatusChanged(ByVal NewStatus As Boolean)
    Public Event ImageChanged(ByVal sender As SignatureBox)
    Public Event FileSet(ByVal sender As SignatureBox, ByVal FileName As String)
    Public Event FileUnset(ByVal sender As SignatureBox)

    Private _FileName As String = Nothing
    Private blnIsSigning As Boolean = False
    Private _blnSigned As Boolean = False
    Private pntLast As Point
    Private Signature As Graphics
    Private BitmapImage As Bitmap
    Public Locked As Boolean = False

    Private Property blnSigned() As Boolean
        Get
            Return Me._blnSigned
        End Get
        Set(ByVal Value As Boolean)
            If Not Me._blnSigned = Value Then
                Me._blnSigned = Value
                RaiseEvent SignedStatusChanged(Value)
            End If
        End Set
    End Property

    Public ReadOnly Property Image() As Bitmap
        Get
            Return Me.BitmapImage
        End Get
    End Property

    Public Property FileName() As String
        Get
            Return _FileName
        End Get
        Set(ByVal Value As String)
            If Value = "" Then
                Value = Nothing
            End If
            If Me._FileName = Value Then
                Exit Property
            Else
                Me._FileName = Value
            End If
            If Not Value = Nothing Then
                If IO.File.Exists(Me._FileName) Then
                    Me.SetFile()
                    Me.blnSigned = True
                    RaiseEvent ImageChanged(Me)
                    Exit Property
                End If
            End If
            Me.NoFile()
        End Set
    End Property

    Property RawData() As String
        Get
            Dim Value As Byte() = Me.GetBytes
            If Value Is Nothing Then
                Return Nothing
            Else
                Return System.Convert.ToBase64String(Value)
            End If
        End Get
        Set(ByVal Value As String)
            If Value.Length > 0 Then
                Dim Mem As New IO.MemoryStream(System.Convert.FromBase64String(Value))
                Dim bmp As Bitmap
                bmp = bmp.FromStream(Mem)
                BitmapImage = New Bitmap(bmp)
                bmp.Dispose()
                Signature = Signature.FromImage(BitmapImage)
                Me.InputArea.Image = BitmapImage
                Me.blnSigned = True
                RaiseEvent ImageChanged(Me)
            Else
                Me.NoFile()
            End If
        End Set
    End Property

    Public Property BorderStyle() As Windows.Forms.BorderStyle
        Get
            Return Me.InputArea.BorderStyle
        End Get
        Set(ByVal Value As Windows.Forms.BorderStyle)
            Me.InputArea.BorderStyle = Value
        End Set
    End Property

    Public ReadOnly Property IsSigned()
        Get
            Return blnSigned
        End Get
    End Property

    Public Property ClearText() As String
        Get
            Return Me.lnkClearSignature.Text
        End Get
        Set(ByVal Value As String)
            Me.lnkClearSignature.Text = Value
        End Set
    End Property

    Public Property ClearTextAlignment() As System.Windows.Forms.HorizontalAlignment
        Get
            If Me.lnkClearSignature.TextAlign = ContentAlignment.MiddleCenter Then
                Return HorizontalAlignment.Center
            ElseIf Me.lnkClearSignature.TextAlign = ContentAlignment.MiddleLeft Then
                Return HorizontalAlignment.Left
            Else
                Return HorizontalAlignment.Right
            End If
        End Get
        Set(ByVal Value As HorizontalAlignment)
            If Value = HorizontalAlignment.Center Then
                Me.lnkClearSignature.TextAlign = ContentAlignment.MiddleCenter
            ElseIf Value = HorizontalAlignment.Left Then
                Me.lnkClearSignature.TextAlign = ContentAlignment.MiddleLeft
            Else
                Me.lnkClearSignature.TextAlign = ContentAlignment.MiddleRight
            End If
        End Set
    End Property

    Public Property ClearTextPosition() As System.Windows.Forms.DockStyle
        Get
            Return Me.lnkClearSignature.Dock
        End Get
        Set(ByVal Value As System.Windows.Forms.DockStyle)
            Me.lnkClearSignature.Dock = Value
        End Set
    End Property

    Public Property ClearTextVisible() As Boolean
        Get
            Return Me.lnkClearSignature.Visible
        End Get
        Set(ByVal Value As Boolean)
            Me.lnkClearSignature.Visible = Value
        End Set
    End Property

    Private Sub SignatureBox_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not Me._FileName = Nothing Then
            If IO.File.Exists(Me._FileName) Then
                Me.SetFile()
                Exit Sub
            End If
        End If
        'Me.NoFile()
    End Sub

    Private Sub SetFile()
        Dim bmp As Bitmap
        bmp = bmp.FromFile(Me._FileName)
        BitmapImage = New Bitmap(bmp)
        bmp.Dispose()
        Signature = Signature.FromImage(BitmapImage)
        Me.InputArea.Image = BitmapImage
        RaiseEvent FileSet(Me, Me._FileName)
        RaiseEvent ImageChanged(Me)
    End Sub

    Private Sub NoFile()
        BitmapImage = New Bitmap(Me.InputArea.Width, Me.InputArea.Height)
        Signature = Signature.FromImage(BitmapImage)
        Signature.Clear(Color.White)
        Me.InputArea.Image = BitmapImage
        Me.blnSigned = False
        RaiseEvent FileUnset(Me)
        RaiseEvent ImageChanged(Me)
    End Sub

    Private Sub Signature_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles InputArea.MouseDown
        blnIsSigning = True
        Me.pntLast = New Point(e.X, e.Y)
        RaiseEvent StartSigning(sender, e)
    End Sub

    Private Sub Signature_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InputArea.Leave
        blnIsSigning = False
        RaiseEvent StopSigning(sender, e)
    End Sub

    Private Sub Signature_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles InputArea.MouseUp
        blnIsSigning = False
        RaiseEvent StopSigning(sender, e)
        RaiseEvent ImageChanged(Me)
    End Sub

    Private Sub Signature_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles InputArea.MouseMove
        If blnIsSigning And Not Locked Then
            Try
                Me.Signature.DrawLine(New Pen(Color.Black, 2), Me.pntLast.X, pntLast.Y, e.X, e.Y)
                Me.InputArea.Image = Me.BitmapImage
                Me.pntLast = New Point(e.X, e.Y)
                Me.blnSigned = True
                RaiseEvent Signing(sender, e)
            Catch
                Me.blnIsSigning = False
                RaiseEvent StopSigning(sender, e)
            End Try
        End If
    End Sub

    Private Sub lnkClearSignature_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkClearSignature.LinkClicked
        Me.ClearInput()
    End Sub

    Public Sub ClearInput()
        BitmapImage = New Bitmap(Me.InputArea.Width, Me.InputArea.Height)
        Signature = Signature.FromImage(BitmapImage)
        Signature.Clear(Color.White)
        Me.InputArea.Image = BitmapImage
        Me.blnSigned = False
        RaiseEvent InputCleared()
        RaiseEvent ImageChanged(Me)
    End Sub

    Public Sub Save(Optional ByVal strFile As String = Nothing)
        If Me.FileName = Nothing And strFile = Nothing Then
            Throw New Exception("No valid file name specified.  Can not save.")
        Else
            If Not strFile = Nothing Then
                Me.FileName = strFile
            End If
            If IO.File.Exists(Me._FileName) Then
                IO.File.Delete(Me._FileName)
            End If
            Me.InputArea.Image.Save(Me._FileName, System.Drawing.Imaging.ImageFormat.Gif)
        End If
    End Sub

    Public Function GetBytes() As Byte()
        Try
            Return Me.BMPToBytes(Me.InputArea.Image)
        Catch
            Return Nothing
        End Try
    End Function

    Private Function BMPToBytes(ByVal bmp As Image) As Byte()
        Dim ms As New System.IO.MemoryStream
        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Gif)
        Dim abyt(ms.Length - 1) As Byte
        ms.Seek(0, IO.SeekOrigin.Begin.Begin)
        ms.Read(abyt, 0, ms.Length)
        Return abyt
    End Function




End Class
