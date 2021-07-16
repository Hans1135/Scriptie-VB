Imports System.IO
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media.Imaging

Module UnitMediaPict

    Public WithEvents MediaPict As New ClassMediaPict

    Class ClassMediaPict
        Inherits Image

        Function Init() As UIElement

            HorizontalAlignment = HorizontalAlignment.Left
            VerticalAlignment = VerticalAlignment.Top
            Stretch = Media.Stretch.Uniform
            Return Me
        End Function

        Sub Done()

            Source = Nothing
        End Sub

        Sub Load(S As String)

            ' laadt een afbeelding

            If File.Exists(S) Then
                MediaDraw.Done()
                MediaScene.Done()
                MediaData.Done()
                MediaScene.Done()

                Dim B As New BitmapImage

                Using F As New FileStream(S, FileMode.Open)
                    B.BeginInit()
                    B.StreamSource = F
                    B.CacheOption = BitmapCacheOption.OnLoad
                    B.EndInit()
                End Using
                Source = B

                If GridMedia.Children.Contains(MediaPict) = False Then GridMedia.Children.Add(MediaPict)
            End If
        End Sub
    End Class
End Module