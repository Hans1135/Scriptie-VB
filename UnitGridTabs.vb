Imports HelixToolkit.Wpf.SharpDX
Imports HelixToolkit.Wpf.SharpDX.Model
Imports HelixToolkit.Wpf.SharpDX.Model.Scene
Imports SharpDX
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media.Media3D

Module UnitGridTabs

    Public GridTabs As New ClassGridTabs

    Class ClassGridTabs
        Inherits TabControl

        Function Init() As UIElement

            Items.Add(TabsCtrl.Init())
            Items.Add(TabsData.Init())
            Items.Add(TabsFile.Init())
            Items.Add(TabsMidi.Init())
            Items.Add(TabsPart.Init())
            Items.Add(TabsPlay.Init())
            Items.Add(TabsPose.Init())
            Items.Add(TabsScene.Init())
            Items.Add(TabsTest.Init())

            TabsFile.IsSelected = True

            Return Me
        End Function
    End Class

    Public TabsData As New ClassTabsData

    Class ClassTabsData
        Inherits TabItem

        Public DataStack As New StackPanel

        Function Init()

            DataStack.Children.Add(TableList)
            Content = DataStack
            Header = "Data"

            Return Me
        End Function
    End Class

    Public TabsPart As New ClassTabsPart

    Class ClassTabsPart
        Inherits TabItem

        Public PartStack As New StackPanel

        Function Init() As TabItem

            PartList.Height = 200

            PartStack.Children.Add(PartList)
            PartStack.Children.Add(AmbientColor)
            PartStack.Children.Add(DiffuseColor)
            PartStack.Children.Add(DiffusePath)
            PartStack.Children.Add(SpecularColor)

            AmbientSlider.Minimum = 0
            AmbientSlider.Value = 0.5
            AmbientSlider.Maximum = 1
            PartStack.Children.Add(AmbientSlider)

            ShininessSlider.Minimum = 0
            ShininessSlider.Value = 0.2
            ShininessSlider.Maximum = 1
            PartStack.Children.Add(ShininessSlider)

            Content = PartStack
            Header = "Part"

            Return Me
        End Function

        Public WithEvents PartList As New ListBox

        Sub PartList_MouseLeftButtonUp() Handles PartList.MouseLeftButtonUp

            ' toont het materiaal van een onderdeel

            Dim I As Short
            Dim M As PhongMaterialCore
            Dim D As DiffuseMaterialCore
            Dim N As MeshNode
            Dim S As String

            S = PartList.SelectedItem
            S = Left(S, InStr(S, ",") - 1)
            I = 0
            For Each Node In SceneModel.SceneNode.Traverse(False)
                If InStr(Node.GetType.ToString, "MeshNode") > 0 Then
                    I += 1
                    If I = S Then
                        N = Node
                        If InStr(LCase(N.Material.GetType.ToString), "phong") > 0 Then
                            M = N.Material
                            AmbientColor.Text = GetColorString(M.AmbientColor)
                            DiffuseColor.Text = GetColorString(M.DiffuseColor)
                            DiffusePath.Text = M.DiffuseMapFilePath
                            SpecularColor.Text = GetColorString(M.SpecularColor)
                            ShininessSlider.Value = M.SpecularShininess
                            Exit Sub
                        Else
                            D = N.Material
                            DiffuseColor.Text = GetColorString(D.DiffuseColor)
                        End If
                    End If
                End If
            Next
        End Sub

        Sub PartList_MouseRightButtonUp() Handles PartList.MouseRightButtonUp

            ' verbergt een onderdeel

            Dim I As Short
            Dim N As MeshNode
            Dim S As String

            S = PartList.SelectedItem
            S = Left(S, InStr(S, ",") - 1)
            I = 0
            For Each Node In SceneModel.SceneNode.Traverse(False)
                If InStr(Node.GetType.ToString, "MeshNode") > 0 Then
                    I += 1
                    If I = S Then
                        N = Node
                        N.Visible = Not N.Visible
                        Exit Sub
                    End If
                End If
            Next
        End Sub

    End Class

    Public TabsPlay As New ClassTabsPlay

    Class ClassTabsPlay
        Inherits TabItem

        Public PlayStack As New StackPanel

        Function Init() As TabItem

            PlayStack.Children.Add(PlayText)
            PositionLabel.Content = "position"
            PlayStack.Children.Add(PositionLabel)
            PlayStack.Children.Add(PlayPosition)
            VolumeLabel.Content = "volume"
            PlayStack.Children.Add(VolumeLabel)
            PlayVolume.Minimum = 0
            PlayVolume.Maximum = 1
            PlayStack.Children.Add(PlayVolume)
            PlayButton.Content = "pause"
            PlayStack.Children.Add(PlayButton)
            PlayStack.Children.Add(PlayMusician)
            PlayStack.Children.Add(PlayYear)
            PlayStack.Children.Add(PlayAlbum)
            PlayStack.Children.Add(PlayNumber)
            PlayStack.Children.Add(PlayTrack)
            PlayStack.Children.Add(PlayCoworker)
            PlayStack.Children.Add(PlayTime)

            Content = PlayStack
            Header = "Play"

            Return Me
        End Function
    End Class

    Public TabsPose As New ClassTabsPose

    Class ClassTabsPose
        Inherits TabItem

        Public PoseStack As New StackPanel

        Function Init()

            BoneList.Height = 200
            BoneList.Items.Clear()

            PoseStack.Children.Add(BoneList)

            SAX.Minimum = -180
            SAX.Maximum = 180
            SAY.Minimum = -180
            SAY.Maximum = 180
            SAZ.Minimum = -180
            SAZ.Maximum = 180

            PoseStack.Children.Add(SDX)
            PoseStack.Children.Add(SDY)
            PoseStack.Children.Add(SDZ)
            PoseStack.Children.Add(SAX)
            PoseStack.Children.Add(SAY)
            PoseStack.Children.Add(SAZ)
            PoseStack.Children.Add(FDX)
            PoseStack.Children.Add(FDY)
            PoseStack.Children.Add(FDZ)
            PoseStack.Children.Add(FAX)
            PoseStack.Children.Add(FAY)
            PoseStack.Children.Add(FAZ)
            PoseStack.Children.Add(TA) ' animatie index
            PoseStack.Children.Add(TB) ' bot index
            PoseStack.Children.Add(TN) ' bot naam
            PoseStack.Children.Add(TK) ' keyframe mode 0,1,2 of 3
            PoseStack.Children.Add(TS) ' synchroon bot
            PoseStack.Children.Add(TX) ' synchroon optie
            PoseStack.Children.Add(TH) ' sliders
            PoseStack.Children.Add(TQ) ' beweeg hoeken als quaternion
            PoseStack.Children.Add(TT) ' keyframe hoek
            Header = "Pose"
            Content = PoseStack

            Return Me
        End Function
    End Class

    Public TabsScene As New ClassTabsScene

    Class ClassTabsScene
        Inherits TabItem

        Function Init() As TabItem

            CameraLens.Minimum = 1
            CameraLens.Maximum = 400
            CameraLens.Value = 50

            ModelLabel.Content = "Model"
            SceneStack.Children.Add(ModelLabel)
            SceneStack.Children.Add(ModelPos)

            CameraLabel.Content = "camera"
            SceneStack.Children.Add(CameraLabel)
            SceneStack.Children.Add(CameraPos)
            SceneStack.Children.Add(LensLabel)
            SceneStack.Children.Add(CameraLens)

            LampLabel.Content = "lampen"
            SceneStack.Children.Add(LampLabel)
            SceneStack.Children.Add(A1Label)
            SceneStack.Children.Add(A1Slider)
            SceneStack.Children.Add(D1Label)
            SceneStack.Children.Add(D1Slider)
            SceneStack.Children.Add(D2Label)
            SceneStack.Children.Add(D2SLider)
            SceneStack.Children.Add(D3Label)
            SceneStack.Children.Add(D3SLider)

            ShadowLabel.Content = "schaduw"
            SceneStack.Children.Add(ShadowLabel)
            SceneStack.Children.Add(OrthowidthLabel)
            'ShadowOrthowidth.Maximum = 1000
            'ShadowOrthowidth.Value = 100
            SceneStack.Children.Add(ShadowOrthowidth)
            SceneStack.Children.Add(IntensityLabel)
            'ShadowIntensity.Maximum = 1
            'ShadowIntensity.Value = 0.1
            SceneStack.Children.Add(ShadowIntensity)
            SceneStack.Children.Add(BiasLabel)
            'ShadowBias.Maximum = 0.004
            'ShadowBias.Value = 0.001
            SceneStack.Children.Add(ShadowBias)
            SceneStack.Children.Add(NearfieldLabel)
            'ShadowNearfield.Maximum = 1
            'ShadowNearfield.Value = 1
            SceneStack.Children.Add(ShadowNearfield)
            SceneStack.Children.Add(FarfieldLabel)
            'ShadowFarfield.Maximum = 500
            'ShadowFarfield.Value = 500
            SceneStack.Children.Add(ShadowFarfield)

            A1Slider.Minimum = 0
            A1Slider.Maximum = 1
            A1Slider.Value = 0.15
            D1Slider.Minimum = 0
            D1Slider.Maximum = 1
            D1Slider.Value = 0.8
            D2SLider.Minimum = 0
            D2SLider.Maximum = 1
            D2SLider.Value = 0.8
            D3SLider.Minimum = 0
            D3SLider.Maximum = 1
            D3SLider.Value = 0.05

            Content = SceneStack
            Header = "Scene"

            Return Me
        End Function

        Sub CameraLens_ValueChanged() Handles CameraLens.ValueChanged

            Dim P As Point3D

            P = SceneCamera.Position
            SceneCamera.Position = New Point3D(P.X + 0.001, P.Y, P.Z) ' truc om camera te verplaatsen anders is er geen update
            LensLabel.Content = "lens " & CameraLens.Value & " mm"
            SceneCamera.FieldOfView = Deg(Math.Atan(36 / (CameraLens.Value * 2)) * 2) ' horizontale beeldhoek van een 50 mm lens
            SceneCamera.Position = P
        End Sub

        Sub A1Slider_ValueChanged() Handles A1Slider.ValueChanged

            A1Label.Content = "lampA1 " & Format(A1Slider.Value, "0.00").ToString
            SceneLamp.LampA1.Color = New Color4(A1Slider.Value, A1Slider.Value, A1Slider.Value, 1).ToColor
        End Sub

        Sub D1Slider_ValueChanged() Handles D1Slider.ValueChanged

            D1Label.Content = "lampD1 " & Format(D1Slider.Value, "0.00").ToString
            SceneLamp.LampD1.Color = New Color4(D1Slider.Value, D1Slider.Value, D1Slider.Value, 1).ToColor
        End Sub

        Sub D2Slider_ValueChanged() Handles D2SLider.ValueChanged

            D2Label.Content = "lampD2 " & Format(D2SLider.Value, "0.00").ToString
            SceneLamp.LampD2.Color = New Color4(D2SLider.Value, D2SLider.Value, D2SLider.Value, 1).ToColor
        End Sub

        Sub D3Slider_ValueChanged() Handles D3SLider.ValueChanged

            D3Label.Content = "lampD3 " & Format(D3SLider.Value, "0.00").ToString
            SceneLamp.LampD3.Color = New Color4(D3SLider.Value, D3SLider.Value, D3SLider.Value, 1).ToColor
        End Sub

        Public CameraLabel As New Label
        Public CameraPos As New TextBox
        Public LensLabel As New Label
        Public WithEvents CameraLens As New Slider
        Public LampLabel As New Label
        Public A1Label As New Label
        Public D1Label As New Label
        Public D2Label As New Label
        Public D3Label As New Label
        Public WithEvents A1Slider As New Slider
        Public WithEvents D1Slider As New Slider
        Public WithEvents D2SLider As New Slider
        Public WithEvents D3SLider As New Slider

    End Class
End Module
