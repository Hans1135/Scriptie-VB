Imports HelixToolkit.Wpf.SharpDX
Imports SharpDX
Imports System.Windows.Controls
Imports System.Windows.Media.Media3D

Module UnitSceneCamera

    Public WithEvents SceneCamera As New ClassSceneCamera

    Class ClassSceneCamera
        Inherits HelixToolkit.Wpf.SharpDX.PerspectiveCamera

        Sub Init()

            ' start de camera

            MediaScene.Camera.Reset
            CA = 0 ' camera hoek met y as
            CF = False
            CS = 1

            Position = New Point3D(-2.4, 1.2, 2.4) ' camera position
            LookDirection = New Vector3D(2.4, -0.6, -2.4) ' camera position
            UpDirection = New Vector3D(0, 1, 0)
            FarPlaneDistance = 1000 '100000 '5000000
            NearPlaneDistance = 0.1 '0.01 ' minimale afstand 10 mm
            FieldOfView = Deg(Math.Atan(36 / (50 * 2)) * 2) ' horizontale beeldhoek van een 50 mm lens
            MediaScene.ZoomSensitivity = 1
            MediaScene.Camera = SceneCamera
            SceneCamera.Show()
        End Sub

        Sub Follow()

            ' beweegt de camera met de ogen van een model mee

            Dim T As Point3D
            Dim D As Vector3D
            Dim M As Matrix
            Dim Q As SharpDX.Quaternion
            Dim E As Vector3
            Dim R As New Transform3DGroup

            If CF Then
                M = AnimUpdt(AI).Animation.NodeAnimationCollection(CH).Node.TotalModelMatrix
                T = M.TranslationVector.ToPoint3D
                T.X += 0.03
                Position = T

                Q = MatrixToQuaternion(M)
                E = QuaternionToEuler(Q)
                ModelPart.Rotate(R, T, -E.X, -E.Y, -E.Z)
                D = New Vector3D(0, -0.8, 2)
                D = R.Transform(D)
                LookDirection = D
                Show()
            End If
        End Sub

        Sub Move()

            ' beweegt de camera met een model mee

            Dim R As New RotateTransform3D
            Dim P As Point3D
            Dim O As Point3D

            If Not CF Then
                R.Rotation = New AxisAngleRotation3D(New Vector3D(0, 1, 0), CA) ' verdraai camera om y-as
                P = R.Transform(New Point3D(0, 1.6 / CS, 3 / CS)) ' verschil positie van camera
                O = ModelAnim(SI).Root.ModelMatrix.TranslationVector.ToPoint3D ' object positie
                Position = O + P ' camera positie
                LookDirection = New Vector3D(-P.X, -0.8 / CS, -P.Z) ' camera richting
            End If
        End Sub

        Sub Pars()

            ' plaatst de camera

            ' W(0) = "camera"
            ' W(1) = type, p of o
            ' V(2) = brandpunt lens in mm
            ' V(3) = zoom factor 
            ' V(4) = X position
            ' V(5) = Y
            ' W(6) = Z
            ' V(7) = X direction
            ' V(8) = Y
            ' V(9) = Z

            If W(1) = "p" Then
                FieldOfView = Deg(Math.Atan(36 / (V(2) * 2)) * 2) ' horizontale beeldhoek
                If W(4) = "follow" Then
                    CH = V(5)
                    CF = True
                Else
                    FarPlaneDistance = V(6) * 40 '2000 '100000
                    NearPlaneDistance = V(6) / 10 ' 1
                    Position = New Point3D(V(4), V(5), V(6))
                    LookDirection = New Vector3D(V(7), V(8), V(9))
                    CS = 2.4 / V(6) ' schaal tov 2.4 m
                End If
                MediaScene.Camera = SceneCamera
                MediaScene.ZoomSensitivity = V(3)
                Show()
            End If
            CA = 0 ' camera hoek
        End Sub

        Sub Show()

            ' toont de camera waardes

            Dim D As Vector3D
            Dim P As Point3D
            Dim F As Single = 18 / (Math.Tan(Rad(SceneCamera.FieldOfView) / 2))

            D = LookDirection
            P = Position
            TabsScene.CameraPos.Text = "camera, p, " & F & ", " & MediaScene.ZoomSensitivity & SF(P.X) & SF(P.Y) & SF(P.Z) & SF(D.X) & SF(D.Y) & SF(D.Z)
        End Sub

        Sub Me_Changed() Handles Me.Changed

            ' als camera verandert

            If SceneLamp.LampFollow Then
                SceneLamp.LampD1.Direction = SceneLamp.Direction()
            End If
            Show()
            If SI > 0 Then
                If ModelAnim(SI) IsNot Nothing Then
                    SceneModelShow()
                End If
            End If
        End Sub

        Public CA As Short ' camera hoek met y-as
        Public CF As Boolean ' camera volgt
        Public CH As Short ' camera volgt hoofd
        Public CS As Short ' camera schaal
    End Class
End Module
