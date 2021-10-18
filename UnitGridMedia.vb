Imports System.IO
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

Module UnitGridMedia

    Public GridMedia As New ClassMedia

    Class ClassMedia
        Inherits Grid

        Function Init() As UIElement

            ' start media veld

            Background = New SolidColorBrush(Colors.White)

            MediaPict.Init()
            MediaPlay.Init()
            MediaDraw.Init()
            MediaScene.Init()
            'MediaData.Init

            Return Me
        End Function

        Sub Done()

            'MediaSceneDone()
        End Sub

        Sub Photo()

            '  maakt een foto van het media veld

            ' W(0) = "photo"
            ' W(1) = [filenaam]

            Dim L As Integer
            Dim H As Integer
            Dim U As UIElement

            If PhotoName = "" Then PhotoName = GetFileAddress("\info.jpg")

            If GridMedia.Children.Contains(MediaScene) Then
                U = MediaScene
            Else
                U = MediaDraw
            End If
            L = U.DesiredSize.Width
            H = U.DesiredSize.Height
            If L = 0 Then Exit Sub
            Dim Bitmap As New RenderTargetBitmap(L, H, 96, 96, PixelFormats.Default)
            Bitmap.Render(U)
            Dim Encoder As New JpegBitmapEncoder()
            Encoder.Frames.Add(BitmapFrame.Create(Bitmap))
            Dim Stream As New FileStream(PhotoName, FileMode.Create)
            Using Stream
                Encoder.Save(Stream)
            End Using
        End Sub
    End Class

    Public PhotoName As String
End Module
