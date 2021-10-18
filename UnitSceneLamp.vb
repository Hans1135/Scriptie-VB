Imports System.Windows.Media.Media3D
Imports HelixToolkit.Wpf.SharpDX
Imports SharpDX

Module UnitSceneLamp

    Public SceneLamp As New ClassSceneLamp

    Class ClassSceneLamp

        Sub Init() ' start de lampen

            MediaScene.Items.Add(LampA1) ' omgeving lamp

            LampD1.Direction = Direction() ' richting lamp 1 schijnt in kijkrichting van de camera
            MediaScene.Items.Add(LampD1)

            LampD2.Direction = New Vector3D(0, -1, 0) ' richting lamp 2 schijnt van boven
            MediaScene.Items.Add(LampD2)

            LampD3.Direction = New Vector3D(1, 0, -1) ' richting lamp 3 schijnt van rechts
            MediaScene.Items.Add(LampD3)

            LampFollow = True ' lamp 1 volgt camera
        End Sub

        Sub Pars()

            ' verandert een lamp

            ' W(0) = "lamp"
            ' W(1) = "a" of "d"
            ' V(2) = rood kleur fraktie
            ' V(3) = groen
            ' V(4) = blauw
            ' V(5) = x richting
            ' V(6) = y
            ' V(7) = z
            ' W(8) = volg

            If W(2) = "" Then V(2) = 1 ' rode kleur
            If W(3) = "" Then V(3) = 1 ' groene kleur
            If W(4) = "" Then V(4) = 1 ' blauwe kleur

            Select Case W(1)
                Case "a1"
                    LampA1.Color = New Color4(V(2), V(3), V(4), 1).ToColor
                Case "d1"
                    LampD1.Color = New Color4(V(2), V(3), V(4), 1).ToColor
                    If W(8) <> "" Then
                        LampFollow = True
                    Else
                        LampFollow = False
                        LampD1.Direction = New Vector3D(V(5), V(6), V(7))
                    End If
                Case "d2"
                    LampD2.Color = New Color4(V(2), V(3), V(4), 1).ToColor
                    LampD2.Direction = New Vector3D(V(5), V(6), V(7))
                Case "d3"
                    LampD3.Color = New Color4(V(2), V(3), V(4), 1).ToColor
                    LampD3.Direction = New Vector3D(V(5), V(6), V(7))
            End Select
        End Sub

        Function Direction() As Vector3D

            Dim D As Vector3D = SceneCamera.LookDirection

            D = New Vector3D(D.X, D.Y * 1.5, D.Z * 1.5)
            Return D
        End Function

        Public LampA1 As New AmbientLight3D
        Public LampD1 As New DirectionalLight3D
        Public LampD2 As New DirectionalLight3D
        Public LampD3 As New DirectionalLight3D
        Public LampFollow As Boolean
    End Class
End Module