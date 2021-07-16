Imports HelixToolkit.Wpf.SharpDX
Imports SharpDX
Imports System.Drawing
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Windows.Media.Media3D

Module UnitMediaScene

    Public MediaScene As New ClassMediaScene

    Class ClassMediaScene
        Inherits Viewport3DX

        Sub Init()

            ' start een omgeving

            SceneCamera.Init()
            SceneLamp.Init()
            SceneModelInit()

            MediaScene.EffectsManager = SceneEffect ' ander blijft scherm zwart ?
            MediaScene.MSAA = MSAALevel.Maximum ' voor nette randen
            MediaScene.IsShadowMappingEnabled = False
            MediaScene.ShowCoordinateSystem = False
            MediaScene.ShowViewCube = False
            MediaScene.InputBindings.Add(MouseLeft)
            MediaScene.InputBindings.Add(MouseRight)
            MediaScene.Items.Add(SceneModel)
        End Sub

        Sub Runs()

            ' start

            SceneCamera.Init()
            SceneModel.Clear()
            SceneModelInit()

            If GridMedia.Children.Contains(MediaScene) = False Then GridMedia.Children.Add(MediaScene)
        End Sub

        Sub Done()

            MediaDraw.Background = New SolidColorBrush(Colors.Transparent)
            GridMedia.Children.Remove(MediaScene)
        End Sub

        Sub Load(S As String)

            Runs()
            W(1) = S
            SceneModelImport()
            If GetFileType(S) <> ".obj" Then ModelAnimStart()
        End Sub

        Sub Floor()

            ' bepaalt de vloer van een omgeving

            ' W(0) = "floor"
            ' V(1) = [grid]
            ' W(2) = [kleur]
            ' W(3) = [bestand]
            ' W(4) = [maak]


            If W(1) = "make" Then
                FloorMap = New Bitmap(256, 256)
                For X = 0 To 255
                    For Y = 0 To 255
                        'FloorMap.SetPixel(X, Y, System.Drawing.Color.White)
                    Next
                Next
                FloorMake = True
            ElseIf W(1) = "save" Then
                FloorMap.Save("d:\data\png\sc\vloer-02.png")
            ElseIf InStr(W(1), ".") > 0 Then
                FloorMap = System.Drawing.Image.FromFile(GetFileAddress(W(3)))
            Else

                Dim F As New AxisPlaneGridModel3D

                F.GridPattern = GridPattern.Tile
                '.GridPattern = GridPattern.Grid
                '.GridThickness = 1
                F.AutoSpacing = False
                F.Offset = V(2)
                F.RenderShadowMap = MediaScene.IsShadowMappingEnabled
                If W(1) <> "" Then F.GridSpacing = V(1)
                If W(2) <> "" Then F.PlaneColor = Colors.Green
                MediaScene.Items.Add(F)
            End If
        End Sub

        Sub Shadow()

            ' voegt schaduw toe

            ' W(0) = "shadow"
            ' V(1) = ?

            MediaScene.IsShadowMappingEnabled = True
            MediaScene.Items.Add(SceneShadow)
            ShadowOrthowidth.Value = V(1)
            ShadowOrthowidth.Maximum = V(1) * 2
            ShadowIntensity.Value = V(2)
            ShadowIntensity.Maximum = V(2) * 2
            ShadowBias.Value = V(3)
            ShadowBias.Maximum = V(3) * 2
            ShadowNearfield.Value = V(4)
            ShadowNearfield.Maximum = V(4) * 2
            ShadowFarfield.Value = V(5)
            ShadowFarfield.Maximum = V(5) * 2
        End Sub

        Sub Spin()

            ' laat de omgeving draaien

            MediaScene.InfiniteSpin = True
            MediaScene.StartSpin(New Windows.Vector(100, 0), New Windows.Point(0, 0), New Point3D(0, 0, 0))
        End Sub

        Sub Me_PreviewMousDown(sender As Object, e As MouseEventArgs) Handles Me.PreviewMouseDown

            If e.MiddleButton = MouseButtonState.Pressed And Paint = True Then
                GridMapsPaint(e.GetPosition(MediaScene).ToVector2)
                e.Handled = True
            ElseIf e.LeftButton = MouseButtonState.Pressed And Game = True Then

                Dim D As Vector3
                Dim P As Vector3
                Dim Q As MeshGeometryModel3D

                Try
                    Q = MediaScene.FindHits(e.GetPosition(MediaScene).ToVector2)(0).ModelHit
                Catch ex As Exception
                    Exit Sub
                End Try

                D = FindNextCard(Q.Name)
                If D.X < 100 Then

                    Dim T As New Transform3DGroup

                    Dim TX As New TranslateTransform3D
                    Dim TY As New TranslateTransform3D
                    Dim TZ As New TranslateTransform3D

                    P.X = Q.TotalModelMatrix.TranslationVector.X
                    P.Y = Q.TotalModelMatrix.TranslationVector.Y
                    P.Z = Q.TotalModelMatrix.TranslationVector.Z
                    Dim AX As New DoubleAnimation(P.X, D.X, New Windows.Duration(TimeSpan.FromSeconds(0.1)))
                    Dim AY As New DoubleAnimation(P.Y, D.Y, New Windows.Duration(TimeSpan.FromSeconds(0.1)))
                    Dim AZ As New DoubleAnimation(P.Z, D.Z + 0.001, New Windows.Duration(TimeSpan.FromSeconds(0.1)))

                    TX.BeginAnimation(TranslateTransform3D.OffsetXProperty, AX)
                    TY.BeginAnimation(TranslateTransform3D.OffsetYProperty, AY)
                    TZ.BeginAnimation(TranslateTransform3D.OffsetZProperty, AZ)
                    T.Children.Add(TX)
                    T.Children.Add(TY)
                    T.Children.Add(TZ)
                    Q.Transform = T
                End If
                e.Handled = True
            End If
        End Sub

        Sub Me_PreviewMouseMove(sender As Object, e As MouseEventArgs) Handles Me.PreviewMouseMove

            FileText.Text = e.GetPosition(MediaScene).ToString
            If e.MiddleButton = MouseButtonState.Pressed And Paint = True Then
                GridMapsPaint(e.GetPosition(MediaScene).ToVector2)
            End If
        End Sub

        Sub Me_PreviewKeyUp() Handles Me.PreviewKeyUp

            KeyPressed = 0
            AnimFileLoad(IdleFile)
        End Sub

        Sub Me_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles Me.PreviewKeyDown

            ' als in het media veld op een toets gedrukt wordt

            If Navigate Then ' als een model genavigeerd kan worden
                If KeyPressed = 0 Then AnimFileLoad(WalkFile)

                KeyPressed = e.Key
                Select Case e.Key
                    Case Key.Up
                        MZ = 1.2 / SceneCamera.CS
                        SceneCamera.CA = MA - 180
                        MT = 1.2
                    Case Key.Down
                        MZ = 1.2 / SceneCamera.CS
                        SceneCamera.CA = MA
                        MT = 1.2
                    Case Key.Left
                        If MA < 360 Then
                            MA += 5
                        Else
                            MA = 0
                        End If
                        SceneCamera.CA = MA - 180
                        MZ = 0
                        MT = 0
                    Case Key.Right
                        If MA > -360 Then
                            MA -= 5
                        Else
                            MA = 0
                        End If
                        SceneCamera.CA = MA - 180
                        MZ = 0
                        MT = 0
                End Select
            End If
            ModelAnimWalk()
            e.Handled = True
            DoEvents()
        End Sub

        Public SceneEffect As New DefaultEffectsManager
        Public MouseLeft = New InputBinding(ViewportCommands.Rotate, New MouseGesture(MouseAction.LeftClick))
        Public MouseRight = New InputBinding(ViewportCommands.Pan, New MouseGesture(MouseAction.RightClick))
        Public KeyPressed As Short
        Public FloorMap As Bitmap
        Public FloorMake As Boolean = False
    End Class

    Sub ShadowIntesity_ValueChanged() Handles ShadowIntensity.ValueChanged

        SceneShadow.Intensity = ShadowIntensity.Value
        IntensityLabel.Content = "intensiteit " & ShadowIntensity.Minimum & " " & ShadowIntensity.Value & " " & ShadowIntensity.Maximum
    End Sub

    Sub ShadowBias_ValueChanged() Handles ShadowBias.ValueChanged

        SceneShadow.Bias = ShadowBias.Value
        BiasLabel.Content = "Bias " & ShadowBias.Minimum & " " & ShadowBias.Value & " " & ShadowBias.Maximum
    End Sub

    Sub ShadowOrthowidth_ValueChanged() Handles ShadowOrthowidth.ValueChanged

        SceneShadow.OrthoWidth = ShadowOrthowidth.Value
        OrthowidthLabel.Content = "Orthowidth " & ShadowOrthowidth.Minimum & " " & ShadowOrthowidth.Value & " " & ShadowOrthowidth.Maximum
    End Sub

    Sub ShadowNearfield_ValueChanged() Handles ShadowNearfield.ValueChanged

        SceneShadow.NearFieldDistance = ShadowNearfield.Value
        NearfieldLabel.Content = "Nearfield " & ShadowNearfield.Minimum & " " & ShadowNearfield.Value & " " & ShadowNearfield.Maximum
    End Sub

    Sub ShadowFarfield_ValueChanged() Handles ShadowFarfield.ValueChanged

        SceneShadow.FarFieldDistance = ShadowFarfield.Value
        FarfieldLabel.Content = "Farfield " & ShadowFarfield.Minimum & " " & ShadowFarfield.Value & " " & ShadowFarfield.Maximum
    End Sub

    Public SceneShadow As New ShadowMap3D
    Public ShadowLabel As New Label
    Public IntensityLabel As New Label
    Public WithEvents ShadowIntensity As New Slider
    Public BiasLabel As New Label
    Public WithEvents ShadowBias As New Slider
    Public OrthowidthLabel As New Label
    Public WithEvents ShadowOrthowidth As New Slider
    Public NearfieldLabel As New Label
    Public WithEvents ShadowNearfield As New Slider
    Public FarfieldLabel As New Label
    Public WithEvents ShadowFarfield As New Slider
    Public SceneStack As New StackPanel
End Module