Namespace GravityDocument


    Public Class gImage

        Dim _Bytes As Byte()
        Dim _Element As gElement

        Public Event ImageChanged(ByVal sender As gImage)

        Public ReadOnly Property Element() As gElement
            Get
                Return Me._Element
            End Get
        End Property

        Public Sub New(ByVal Parent As gElement)
            Me._Element = Parent
        End Sub

        Private Function BMPToBytes(ByVal bmp As System.Drawing.Image) As Byte()
            Dim ms As New System.IO.MemoryStream
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Gif)
            Dim abyt(ms.Length - 1) As Byte
            ms.Seek(0, IO.SeekOrigin.Begin)
            ms.Read(abyt, 0, ms.Length)
            Return abyt
        End Function

        Public Sub Clear(ByVal bg As System.Drawing.Color)
            Dim bmp As New Drawing.Bitmap(Me.Element.Width, Me.Element.Height)
            Dim g As Drawing.Graphics = Drawing.Graphics.FromImage(bmp)
            g.Clear(bg)
            Me.LoadFromBmp(bmp)
            RaiseEvent ImageChanged(Me)
        End Sub

        Public Sub LoadFromBmp(ByVal bmp As System.Drawing.Image)
            Me._Bytes = Me.BMPToBytes(bmp)
            RaiseEvent ImageChanged(Me)
        End Sub

        Public Sub LoadFromDataUrl(ByVal Src As String)
            If Src.StartsWith("data:image/gif;base64,") Then
                Dim strBytes As String = Src.Substring(22)
                Me._Bytes = System.Convert.FromBase64String(strBytes)
                RaiseEvent ImageChanged(Me)
            End If
        End Sub

        Public Sub LoadFromString(ByVal str As String)
            Me._Bytes = System.Convert.FromBase64String(str)
        End Sub

        Public Function ToMemoryStream() As System.IO.MemoryStream
            If Me._Bytes IsNot Nothing Then
                Return New System.IO.MemoryStream(Me._Bytes)
            Else
                Return Nothing
            End If
        End Function

        Public Function ToDataUrl() As String
            If Me._Bytes IsNot Nothing Then
                Return "data:image/gif;base64," & System.Convert.ToBase64String(Me._Bytes)
            Else
                Return Nothing
            End If
        End Function

        Public Overrides Function ToString() As String
            If Me._Bytes IsNot Nothing Then
                Return System.Convert.ToBase64String(Me._Bytes)
            Else
                Return Nothing
            End If
        End Function

        Public Function ToBmp() As System.Drawing.Bitmap
            If Me._Bytes IsNot Nothing Then
                Return New System.Drawing.Bitmap(Me.ToMemoryStream)
            Else
                Return Nothing
            End If
        End Function

    End Class

End Namespace
