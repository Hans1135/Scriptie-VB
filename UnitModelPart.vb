Imports HelixToolkit.Wpf.SharpDX
Imports HelixToolkit.Wpf.SharpDX.Model.Scene
Imports SharpDX
Imports System.Data
Imports System.IO
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Windows.Media.Media3D

Module UnitModelPart

    Public ModelPart As New ClassModelPart

    Class ClassModelPart
        Inherits MeshGeometryModel3D

        Sub Init()

            ' start onderdeel

            PartMaterial.Init()
            PartLegoInit()
            PartMove.T = 0
            PartTurn.T = 0
            PartKeep = False
        End Sub

        Sub Add()

            ' voegt een onderdeel aan een model toe

            Dim B As BoundingBox
            Dim C As System.Drawing.Color
            Dim C2 As System.Drawing.Color
            Dim D As Vector3
            Dim G As HelixToolkit.Wpf.SharpDX.MeshGeometry3D
            Dim H As Single
            Dim N As Long
            Dim T As New Transform3DGroup

            If ModelPart IsNot Nothing Then
                If ModelPart.Geometry IsNot Nothing Then
                    If ModelPart.Geometry.Positions.Count > 0 Then
                        T.Children.Add(ModelPart.Transform)
                        T.Children.Add(PartMoveGet)
                        T.Children.Add(PartTurnGet)
                        ModelPart.Transform = T
                        If UVCorr(1).X > 0 Then
                            G = ModelPart.Geometry
                            N = G.TextureCoordinates.Count
                            For I = 0 To N - 1
                                G.TextureCoordinates(I) *= UVCorr(1)
                            Next
                        End If
                        'G = ModelPart.Geometry
                        'N = G.Normals.Count
                        'For I = 0 To N - 1
                        'G.Normals(I) = New Vector3(1, 1, 1)
                        'Next
                        ModelPart.IsThrowingShadow = MediaScene.IsShadowMappingEnabled
                        SceneModel.AddNode(ModelPart)
                        If MediaScene.FloorMake Then
                            N = SceneModel.SceneNode.Items.Count - 1 ' laatste onderdeel
                            B = SceneModel.SceneNode.Items(N).Bounds
                            D = SceneModel.SceneNode.Items(N).ModelMatrix.TranslationVector
                            H = (D.Y + B.Maximum.Y) / 0.0016
                            If H > 100 Then
                                'MsgBox(H & " " & ParsFile(FI - 1))
                            End If
                            C = System.Drawing.Color.FromArgb(255, H, H, H)
                            For X = 127 + (D.X + B.Minimum.X) / 0.008 To 127 + (D.X + B.Maximum.X) / 0.008
                                For Y = 127 + (D.Z + B.Minimum.Z) / 0.008 To 127 + (D.Z + B.Maximum.Z) / 0.008
                                    C2 = MediaScene.FloorMap.GetPixel(X, Y)
                                    If C2.G = 0 Then
                                        MediaScene.FloorMap.SetPixel(X, Y, C)
                                    End If
                                Next
                            Next
                        End If
                    End If
                End If
                If PartMove.T > 0 Then Wait(PartMove.T)
                ModelPart = New ClassModelPart
            End If
        End Sub

        Sub Box()

            ' tekent een balk

            ' W(0) = "box"
            ' V(1) = x breedte
            ' V(2) = y hoogte
            ' V(3) = z diepte
            ' V(4) = x offset
            ' V(5) = y
            ' V(6) = z
            ' V(7) = x hoek om middelpunt
            ' V(8) = y
            ' V(9) = z
            ' V(10) = x schaal
            ' V(11) = y
            ' V(12) = z
            ' W(13) = naam

            Dim B As New MeshBuilder
            Dim T As New Transform3DGroup

            B.AddBox(New Vector3(0, 0, 0), V(1), V(2), V(3))
            Geometry = B.ToMeshGeometry3D()
            Material = PartMaterial

            FI += 1
            OX(FI) = V(4)
            OY(FI) = V(5)
            OZ(FI) = V(6)
            AX(FI) = V(7)
            AY(FI) = V(8)
            AZ(FI) = V(9)
            If W(10) <> "" Then SX(FI) = V(10) Else SX(FI) = 1
            If W(11) <> "" Then SY(FI) = V(11) Else SY(FI) = 1
            If W(12) <> "" Then SZ(FI) = V(12) Else SZ(FI) = 1

            If UVCorr(0).X > 0 Then
                UVCorr(1).X = V(1) * SX(FI) / UVCorr(0).X
                UVCorr(1).Y = V(2) * SY(FI) / UVCorr(0).Y
            Else
                UVCorr(1).X = 1
                UVCorr(1).Y = 1
            End If

            For I = FI To 0 Step -1
                Scale(T, New Vector3D(SX(I), SY(I), SZ(I)))
                Translate(T, OX(I), OY(I), OZ(I))
                Rotate(T, New Point3D(OX(I), OY(I), OZ(I)), AX(I), AY(I), AZ(I))
            Next
            FI -= 1
            Transform = T
            Name = W(13)
            Add()
        End Sub

        Sub Cylinder()

            ' tekent een cilinder

            ' W(0) = "cylinder"
            ' V(1) = diameter
            ' V(2) = hoogte
            ' V(3) = segmenten
            ' V(4) = dx
            ' V(5) = dy
            ' V(6) = dz
            ' V(7) = hx
            ' V(8) = hy
            ' V(9) = hz

            Dim B As New MeshBuilder
            Dim T As New Transform3DGroup

            ' onderkant, bovenkant, straal, segmenten, bodem, deksel
            B.AddCylinder(New SharpDX.Vector3(0, -V(2) / 2, 0), New SharpDX.Vector3(0, V(2) / 2, 0), V(1) / 2, V(3), True, True)
            Geometry = B.ToMeshGeometry3D()
            Material = PartMaterial

            FI += 1
            OX(FI) = V(4)
            OY(FI) = V(5)
            OZ(FI) = V(6)
            AX(FI) = V(7)
            AY(FI) = V(8)
            AZ(FI) = V(9)

            For I = FI To 0 Step -1
                Scale(T, New Vector3D(SX(I), SY(I), SZ(I)))
                Translate(T, OX(I), OY(I), OZ(I))
                Rotate(T, New Point3D(OX(I), OY(I), OZ(I)), AX(I), AY(I), AZ(I))
            Next

            FI -= 1
            Transform = T
            Add()
        End Sub

        Sub Pyramid()

            ' teken een piramide

            ' W(0) = "pyramid"
            ' V(1) = 

            Dim B As New MeshBuilder
            Dim T As New Transform3DGroup

            B.AddPyramid(New SharpDX.Vector3(0, V(2) / 3, 0), New SharpDX.Vector3(1, 0, 0), New SharpDX.Vector3(0, 1, 0), V(1), V(2), True)
            Geometry = B.ToMeshGeometry3D()
            Material = PartMaterial

            FI += 1
            OX(FI) = V(3)
            OY(FI) = V(4)
            OZ(FI) = V(5)
            AX(FI) = V(6)
            AY(FI) = V(7)
            AZ(FI) = V(8)
            For I = FI To 0 Step -1
                Scale(T, New Vector3D(SX(I), SY(I), SZ(I)))
                Translate(T, OX(I), OY(I), OZ(I))
                Rotate(T, New Point3D(OX(I), OY(I), OZ(I)), AX(I), AY(I), AZ(I))
            Next
            FI -= 1

            Transform = T
            Add()
        End Sub
        '
        Sub Segment()

            ' tekent een segment

            ' W(0) = "segment"
            ' V(1) = buiten straal
            ' V(2) = binnen straal
            ' V(3) = hoogte
            ' V(4) = segmenten voor 360°
            ' V(5) = hoek
            ' V(6) = dx
            ' V(7) = dy
            ' V(8) = dz
            ' V(10) = hx
            ' V(11) = hy
            ' V(12) = hz

            Dim B As New MeshBuilder
            Dim N As Short

            Dim P(200) As SharpDX.Vector3
            Dim T As New Transform3DGroup

            V(3) /= 2 ' hoogte
            If V(4) = 0 Then V(4) = 24 ' aantal segmenten in 360°
            V(4) = 360 / V(4) ' graden per segment
            If V(5) = 0 Then V(5) = 360 ' hoek
            N = Math.Abs(V(5) / V(4)) ' aantal segmenten voor hoek
            V(4) *= -1
            For I = 0 To N ' bereken punten vanaf 12 uur met klok mee
                P(I + 0) = New SharpDX.Vector3(DXC(V(2), I * V(4) + 90), V(3), -DYS(V(2), I * V(4) + 90)) ' bovenkant binnen
                P(I + 40) = New SharpDX.Vector3(DXC(V(2), I * V(4) + 90), -V(3), -DYS(V(2), I * V(4) + 90)) ' onderkant binnen
                P(I + 80) = New SharpDX.Vector3(DXC(V(1), I * V(4) + 90), -V(3), -DYS(V(1), I * V(4) + 90)) ' onderkant buiten
                P(I + 120) = New SharpDX.Vector3(DXC(V(1), I * V(4) + 90), V(3), -DYS(V(1), I * V(4) + 90)) ' bovenkant buiten
            Next

            'B.CreateNormals = True
            For I = 0 To N - 1
                B.AddQuad(P(I + 120), P(I), P(I + 1), P(I + 121)) ' bovenkant
                B.AddQuad(P(I + 81), P(I + 41), P(I + 40), P(I + 80)) ' onderkant
                B.AddQuad(P(I + 121), P(I + 81), P(I + 80), P(I + 120)) ' buitenkant
                B.AddQuad(P(I + 0), P(I + 40), P(I + 41), P(I + 1)) ' binnenkant
            Next
            B.AddQuad(P(120), P(80), P(40), P(0)) ' voorkant
            B.AddQuad(P(N), P(N + 40), P(N + 80), P(N + 120)) ' achterkant
            Geometry = B.ToMeshGeometry3D()
            Material = PartMaterial

            FI += 1
            OX(FI) = V(6)
            OY(FI) = V(7)
            OZ(FI) = V(8)
            AX(FI) = V(9)
            AY(FI) = V(10)
            AZ(FI) = V(11)
            For I = FI To 0 Step -1
                Translate(T, OX(I), OY(I), OZ(I))
                Rotate(T, New Point3D(OX(I), OY(I), OZ(I)), AX(I), AY(I), AZ(I))
            Next
            FI -= 1
            Transform = T
            Add()
        End Sub

        Sub Sphere()

            ' teken een bol

            ' W(0) = "sphere"
            ' V(1) = diameter
            ' V(2) = [segmenten]
            ' V(3) = [X center]
            ' V(4) = [Y center]
            ' V(5) = [Z center]

            Dim B As New MeshBuilder
            Dim T As New Transform3DGroup

            If W(2) = "" Then V(2) = 32

            B.AddSphere(New Vector3(0, 0, 0), V(1) / 2, V(2), V(2))
            Geometry = B.ToMeshGeometry3D()
            Material = PartMaterial
            FI += 1
            OX(FI) = V(3)
            OY(FI) = V(4)
            OZ(FI) = V(5)
            AX(FI) = V(6)
            AY(FI) = V(7)
            AZ(FI) = V(8)
            If W(9) <> "" Then SX(FI) = V(9)
            If W(10) <> "" Then SY(FI) = V(10)
            If W(11) <> "" Then SZ(FI) = V(11)
            For I = FI To 0 Step -1
                Scale(T, New Vector3D(SX(I), SY(I), SZ(I)))
                Translate(T, OX(I), OY(I), OZ(I))
                Rotate(T, New Point3D(OX(I), OY(I), OZ(I)), AX(I), AY(I), AZ(I))
            Next
            FI -= 1
            Transform = T
            Add()
        End Sub

        Sub Slope()
            '
            ' teken een helling
            '
            ' W(0) = "slope"
            ' V(1) = X breedte
            ' V(2) = Y hoogte
            ' V(3) = Z diepte
            ' V(4) = X zwaartepunt
            ' V(5) = Y zwaartepunt
            ' V(6) = Z zwaartepunt
            ' V(7) = HX
            ' V(8) = HY
            ' V(9) = HZ
            '
            Dim PD As New Point3D()
            Dim PV(6) As SharpDX.Vector3
            Dim B As New MeshBuilder
            Dim T As New Transform3DGroup

            For I = 1 To 3
                V(I) /= 2
            Next

            PV(1).X = -V(1) ' loa
            PV(1).Y = -V(2)
            PV(1).Z = -V(3)
            PV(2).X = -V(1) ' lov
            PV(2).Y = -V(2)
            PV(2).Z = V(3)
            PV(3).X = V(1) ' rov
            PV(3).Y = -V(2)
            PV(3).Z = V(3)
            PV(4).X = V(1) ' roa
            PV(4).Y = -V(2)
            PV(4).Z = -V(3)

            PV(5).X = -V(1) ' lba
            PV(5).Y = V(2)
            PV(5).Z = -V(3)
            PV(6).X = V(1) ' rba
            PV(6).Y = V(2)
            PV(6).Z = -V(3)

            B.AddQuad(PV(2), PV(1), PV(4), PV(3)) ' grondvlak
            B.AddQuad(PV(6), PV(4), PV(1), PV(5)) ' achtervlak
            B.AddQuad(PV(5), PV(2), PV(3), PV(6)) ' schuine vlak
            B.AddTriangle(PV(1), PV(2), PV(5)) ' links
            B.AddTriangle(PV(3), PV(4), PV(6)) ' rechts
            Geometry = B.ToMeshGeometry3D()
            Material = PartMaterial
            FI += 1
            OX(FI) = V(4)
            OY(FI) = V(5)
            OZ(FI) = V(6)
            AX(FI) = V(7)
            AY(FI) = V(8)
            AZ(FI) = V(9)
            For I = FI To 0 Step -1
                Scale(T, New Vector3D(SX(I), SY(I), SZ(I)))
                Translate(T, OX(I), OY(I), OZ(I))
                Rotate(T, New Point3D(OX(I), OY(I), OZ(I)), AX(I), AY(I), AZ(I))
            Next
            FI -= 1
            Transform = T
            Add()
        End Sub

        Sub Rotate(T As Media3D.Transform3DGroup, P As Media3D.Point3D, DX As Single, DY As Single, DZ As Single)

            ' draait een onderdeel

            T.Children.Add(New Media3D.RotateTransform3D(New Media3D.AxisAngleRotation3D(New Media3D.Vector3D(1, 0, 0), DX), P))
            T.Children.Add(New Media3D.RotateTransform3D(New Media3D.AxisAngleRotation3D(New Media3D.Vector3D(0, 1, 0), DY), P))
            T.Children.Add(New Media3D.RotateTransform3D(New Media3D.AxisAngleRotation3D(New Media3D.Vector3D(0, 0, 1), DZ), P))
        End Sub

        Sub Scale(T As Media3D.Transform3DGroup, P As Media3D.Vector3D)

            ' schaalt een onderdeel

            T.Children.Add(New Media3D.ScaleTransform3D(P))
        End Sub

        Sub Translate(T As Media3D.Transform3DGroup, X As Double, Y As Double, z As Double)

            ' verplaatst een onderdeel

            T.Children.Add(New Media3D.TranslateTransform3D(X, Y, z))
        End Sub
    End Class

    '
    Sub ModelPartFace()

        ' teken een vlak

        ' W(0) = "face"of "f"
        ' V(1) = punt 1 als p of p//n of p/t/n
        ' V(2) = punt 2
        ' V(3) = punt 3
        ' V(4) = punt 4

        Dim B As New MeshBuilder
        Dim I As Short
        Dim J As Short
        Dim K As Short
        Dim N As Long
        Dim T As New Transform3DGroup

        If PartKeep = False Then
            PartBuilder = New MeshBuilder
            ModelPartNew()
        End If

        I = InStr(W(1), "//")
        If I > 0 Then
            For J = 1 To 4
                If W(J) <> "" Then
                    I = InStr(W(J), "//")
                    V(J) = Val(Left(W(J), I - 1)) ' punt index
                    If V(J) < 0 Then V(J) = PN - V(J) + 1
                    V(J + 8) = Val(Mid(W(J), I + 2)) ' normaal index
                End If
            Next
        Else
            I = InStr(W(1), "/")
            If I > 0 Then
                For J = 1 To 4
                    If W(J) <> "" Then
                        I = InStr(W(J), "/") ' eerste slash
                        K = InStrRev(W(J), "/") ' tweede slash
                        V(J) = Val(Left(W(J), I - 1))
                        If V(J) < 0 Then V(J) = PN - V(J) + 1
                        If K = I Then
                            V(J + 4) = Val(Mid(W(J), I + 1))
                            V(J + 8) = 0
                        Else
                            V(J + 4) = Val(Mid(W(J), I + 1, K - I - 1))
                            V(J + 8) = Val(Mid(W(J), K + 1))
                        End If
                    End If
                Next
            End If
        End If

        If W(4) = "" Then ' zonder punt 4 wordt een driehoek getekend
            PartBuilder.AddTriangle(PA(V(1)), PA(V(2)), PA(V(3)), UA(V(5)), UA(V(6)), UA(V(7)))
        Else ' met een 4e punt wordt een vierhoek getekend
            PartBuilder.AddQuad(PA(V(1)), PA(V(2)), PA(V(3)), PA(V(4)), UA(V(5)), UA(V(6)), UA(V(7)), UA(V(8)))
        End If
        If NN > 0 Then
            Dim G As HelixToolkit.Wpf.SharpDX.MeshGeometry3D

            G = PartBuilder.ToMeshGeometry3D()
            N = G.Normals.Count
            If W(4) = "" Then
                G.Normals(N - 3) = NA(V(9))
                G.Normals(N - 2) = NA(V(10))
                G.Normals(N - 1) = NA(V(11))
            Else
                G.Normals(N - 4) = NA(V(9))
                G.Normals(N - 3) = NA(V(10))
                G.Normals(N - 2) = NA(V(11))
                G.Normals(N - 1) = NA(V(12))
            End If
        End If
        If PartKeep = False Then
            ModelPart.Geometry = PartBuilder.ToMeshGeometry3D()
            ModelPart.Material = PartMaterial

            If UVCorr(0).X <> 0 Then ' ???
                UVCorr(1).X = (PA(3).X - PA(1).X) * SX(FI) / UVCorr(0).X
                UVCorr(1).Y = (PA(1).Y - PA(2).Y) * SY(FI) / UVCorr(0).Y
            Else
                UVCorr(1).X = 1
                UVCorr(1).Y = 1
            End If

            For I = FI To 0 Step -1
                ModelPart.Scale(T, New Vector3D(SX(I), SY(I), SZ(I)))
                ModelPart.Translate(T, OX(I), OY(I), OZ(I))
                ModelPart.Rotate(T, New Point3D(OX(I), OY(I), OZ(I)), AX(I), AY(I), AZ(I))
            Next
            ModelPart.Transform = T
            ModelPart.Add()
            'PartBuilder = New MeshBuilder
        End If
    End Sub

    Sub ModelPartKeep()

        ' W(0) = "keep"

        If PartKeep = False Then
            PartBuilder.Positions.Clear()
            ModelPartNew()
            PartKeep = True
        Else
            ModelPart.Geometry = PartBuilder.ToMeshGeometry3D
            ModelPart.Material = PartMaterial
            ModelPart.Add()
            PartKeep = False
        End If
    End Sub

    Sub ModelPartLetters()
        ' 
        ' maak letter objecten zoals in helix demo
        '
    End Sub
    '
    Sub ModelPartMap()
        '
        ' importeert textuur coordinaten
        '
        ' W(0) = "map"
        ' V(1) = index
        ' V(2) = X
        ' V(3) = Y
        '
        Dim P As Vector2
        '
        PartGeo = ModelPart.Geometry
        P.X = V(2)
        P.Y = V(3)
        PartGeo.TextureCoordinates(V(1)) = P
    End Sub
    '
    Sub ModelPartMaps()
        '
        ' exporteert textuurcoordinaten
        '
        ' W(0) = "maps"
        ' W(1) = filenaam
        '
        Dim N As Short
        Dim S As String = ""
        '
        PartGeo = ModelPart.Geometry
        N = PartGeo.TextureCoordinates.Count - 1
        For I = 0 To N
            S &= "map, " & I & ", " & PartGeo.TextureCoordinates(I).X & ", " & PartGeo.TextureCoordinates(I).Y & ", " &
            PartGeo.Positions(I).X & ", " & PartGeo.Positions(I).Y & ", " & PartGeo.Positions(I).Z & ", " & vbCrLf
        Next
        File.WriteAllText(GetFileAddress(W(1)), S)
    End Sub

    Sub ModelPartMove()
        '
        ' verplaatst een object
        '
        ' W(0) = "move"
        ' V(1) = X
        ' V(2) = Y
        ' V(3) = Z
        ' V(4) = tijd in secondes
        ' V(5) = herhalen, 0 = niet, 1 = wel, 2 =  heen en weer
        '
        PartMove.X = V(1)
        PartMove.Y = V(2)
        PartMove.Z = V(3)
        PartMove.T = V(4)
        PartMove.R = V(5)
    End Sub

    Sub ModelPartNew()

        ' maakt een nieuw onderdeel

        ModelPart = New ClassModelPart
        'ModelPart.Name = "test"
    End Sub

    Sub ModelPartNormal()

        ' bewaart een normaal

        ' W(0) = "normaal", "normal"
        ' V(1) = index
        ' V(2) = X
        ' V(3) = Y
        ' V(4) = Z

        If W(4) <> "" Then ' met index
            NN = V(1)
            NA(NN).X = V(2)
            NA(NN).Y = V(3)
            NA(NN).Z = V(4)
        Else ' zonder index
            NN += 1
            NA(NN).X = V(1)
            NA(NN).Y = V(2)
            NA(NN).Z = V(3)
        End If
    End Sub

    Sub ModelPartPillar()

        ' tekent een pilaar

        ' W(0) = "pillar"
        ' W(1) = X
        ' W(2) = Y
        ' W(3) = U

        Dim B As New MeshBuilder
        Dim D As Vector3
        Dim O As Vector3
        Dim P As New Collections.Generic.List(Of Vector2)
        Dim S As Short
        Dim T As New Transform3DGroup
        Dim U As New Collections.Generic.List(Of Double)

        ModelPartNew()

        For I = 0 To PN
            P.Add(New Vector2(PA(I).X, PA(I).Y))
            U.Add(PA(I).Z)
        Next

        O = New Vector3(0, 0, 0)
        D = New Vector3(0, 1, 0)
        S = 25 ' geeft 24 segmenten

        If W(3) = "" Then
            B.AddRevolvedGeometry(P, U, O, D, S)
        Else
            B.AddRevolvedGeometry(P, Nothing, O, D, S)
        End If

        ModelPart.Geometry = B.ToMeshGeometry3D()
        ModelPart.Material = PartMaterial

        FI += 1
        OX(FI) = V(4)
        OY(FI) = V(5)
        OZ(FI) = V(6)
        AX(FI) = V(7)
        AY(FI) = V(8)
        AZ(FI) = V(9)
        If W(10) <> "" Then SX(FI) = V(10) Else SX(FI) = 1
        If W(11) <> "" Then SY(FI) = V(11) Else SY(FI) = 1
        If W(12) <> "" Then SZ(FI) = V(12) Else SZ(FI) = 1

        For I = FI To 0 Step -1
            ModelPart.Scale(T, New Vector3D(SX(I), SY(I), SZ(I)))
            ModelPart.Translate(T, OX(I), OY(I), OZ(I))
            ModelPart.Rotate(T, New Point3D(OX(I), OY(I), OZ(I)), AX(I), AY(I), AZ(I))
        Next
        FI -= 1
        ModelPart.Transform = T
        ModelPart.Add()
    End Sub

    Sub ModelPartPoint()

        ' bewaart een punt

        ' W(0) = "point"
        ' V(1) = index
        ' V(2) = X
        ' V(3) = Y
        ' V(4) = Z

        If W(4) <> "" Then ' punt heeft index
            PA(V(1)).X = V(2)
            PA(V(1)).Y = V(3)
            PA(V(1)).Z = V(4)
        Else
            PN += 1 ' punt heeft geen index
            PA(PN).X = V(1)
            PA(PN).Y = V(2)
            PA(PN).Z = V(3)
        End If
    End Sub

    Sub ModelPartPoints()

        ' exporteert de punten van een onderdeel

        ' W(0) = "points"
        ' W(1) = bestandnaam

        Dim I As Long
        Dim J As Long
        Dim K As Long
        Dim L As Long
        Dim P As HelixToolkit.Wpf.SharpDX.MeshGeometry3D
        Dim N As Long
        Dim M As MeshNode
        Dim S As String
        Dim X(2) As Single
        Dim Y(2) As Single
        Dim Z(2) As Single

        W(1) = GetFileAddress(W(1))
        FileOpen(1, W(1), OpenMode.Output)
        PMax = New Vector3(V(2), V(3), V(4))
        PMin = New Vector3(V(5), V(6), V(7))

        For Each Node In SceneModel.SceneNode.Traverse(True)
            S = Node.GetType.ToString
            If InStr(LCase(S), "meshnode") > 0 Then
                M = Node
                P = M.Geometry
                N = P.TriangleIndices.Count - 1
                For I = 0 To N Step 3 ' iedere driehoek heeft 3 posities
                    For J = 0 To 2
                        K = P.TriangleIndices(I + J)
                        X(J) = P.Positions((K)).X
                        If PMax.X < X(J) Or PMin.X > X(J) Then GoTo volgende
                        Y(J) = P.Positions((K)).Y
                        If PMax.Y < Y(J) Or PMin.Y > Y(J) Then GoTo volgende
                        Z(J) = P.Positions((K)).Z
                        If PMax.Z < Z(J) Or PMin.Z > Z(J) Then GoTo volgende
                    Next
                    For J = 0 To 2
                        PrintLine(1, "v, " & J + 1 & SF(X(J)) & SF(Y(J)) & SF(Z(J)))
                    Next
                    For J = 0 To 2
                        K = P.TriangleIndices(I + J)
                        PrintLine(1, "vn, " & J + 1 & SF(P.Normals(K).X) & SF(P.Normals(K).Y) & SF(P.Normals(K).Z))
                    Next
                    PrintLine(1, "f, 1//1, 2//2, 3//3")
                    L += 1
volgende:
                Next
                PrintLine(1, "' " & L)
                FileClose(1)
                Exit Sub
            End If
        Next
    End Sub
    '
    Sub ModelPartTexture()
        '
        ' texture coordinaten opslaan
        '
        ' W(0) = "texture" of "vt"
        ' V(1) = index
        ' V(2) = U
        ' V(3) = V
        ' V(4) = [w] ' bij veelvouden worden meerdere textuur files naast en onder elkaar gebruikt
        ' V(5) = [h]
        '
        'If W(0) = "vt" Then 
        'UN += 1
        'UA(UN).X = V(1)
        'UA(UN).Y = 1 - V(2)
        'Else
        UA(V(1)).X = V(2)
        UA(V(1)).Y = V(3)
        'End If
        If W(4) <> "" Then
            UA(V(1)).X *= V(4)
            UA(V(1)).Y *= V(5)
        End If
    End Sub

    Sub ModelPartExport()

        ' maakt een lijst met onderdelen van een model

        Dim I As Short
        Dim M As Model.MaterialCore
        Dim N As MeshNode

        If W(1) <> "" Then
            W(1) = GetFileAddress(W(1))
            FileOpen(1, W(1), OpenMode.Output)
        End If

        I = 0
        TabsPart.PartList.Items.Clear()

        For Each Node In SceneModel.SceneNode.Traverse(False)
            If InStr(Node.GetType.ToString, "MeshNode") > 0 Then
                I += 1
                N = Node
                M = N.Material
                TabsPart.PartList.Items.Add(I & ", " & N.Name & ", " & M.Name)
                If W(1) <> "" Then PrintLine(1, "part, " & I & ", " & N.Name.ToString & ", ")
            End If
        Next
        FileClose(1)
    End Sub

    Sub ModelPartImport()

        ' importeert een onderdelen, verbergt eventueel, hernoemt eventueel

        ' W(0) = "part"
        ' V(1)..V(20) = [nummer van onderdelen]
        ' W(2) = [naam]
        ' W(3) = [nieuwe naam]

        Dim I As Short
        'Dim J As Short
        Dim N As MeshNode

        I = 0
        If ModelAnim(SI) IsNot Nothing Then
            For Each Node In ModelAnim(SI).Root.Traverse(False)
                If InStr(Node.GetType.ToString, "MeshNode") > 0 Then
                    I += 1
                    N = Node

                    If I = V(1) Then
                        If W(3) <> "" Then ' hernoemt onderdeel
                            'MsgBox(N.Name & " hernoemen in " & W(3))
                            N.Name = W(3)
                        End If
                    End If
                    For J = 1 To 20 ' verbergt onderdeel
                        If W(J) <> "" Then
                            If I = V(J) Then N.Visible = False
                        End If
                    Next
                End If
            Next
        Else
            For Each Node In SceneModel.SceneNode.Traverse(False)
                If InStr(Node.GetType.ToString, "MeshNode") > 0 Then
                    I += 1
                    N = Node
                    If I = V(1) Then
                        If W(3) <> "" Then ' hernoemt onderdeel
                            N.Name = W(3)
                        End If
                    End If
                    'For J = 1 To 20 ' verbergt onderdeel
                    'If W(J) <> "" Then
                    'If I = V(J) Then N.Visible = False
                    'End If
                    'Next
                End If
            Next
        End If
    End Sub

    Sub ModelPartReform()

        ' hervormt een model om geexporteerd te kunnen worden

        ' W(0) = "reform"
        ' V(1) = type, 0 = model, 1 = part
        ' V(2) = [schaal]

        Dim I As Long
        Dim J As Long
        Dim K As Long = 0
        Dim M As Matrix3D
        Dim N As MeshNode
        Dim G As HelixToolkit.Wpf.SharpDX.MeshGeometry3D
        Dim O As Media3D.Point3D
        Dim P As Media3D.Point3D
        Dim T As New MatrixTransform3D()
        Dim S As New ScaleTransform3D(New Vector3D(SC, SC, SC))

        If V(1) = 0 And SI = 0 Then V(1) = 1
        If V(1) = 0 Then
            For Each Node In ModelAnim(SI).Root.Traverse(True)
                'MsgBox(Node.Name & ", " & Node.ModelMatrix.TranslationVector.ToString)
                If InStr(LCase(Node.GetType.Name), "meshnode") > 0 Then
                    N = Node
                    'MsgBox(N.Name)
                    M = N.ModelMatrix.ToMatrix3D
                    T.Matrix = M * S.ToMatrix.ToMatrix3D
                    G = N.Geometry
                    J = G.Positions.Count - 1
                    For I = 0 To J
                        O = G.Positions(I).ToPoint3D
                        P = T.Transform(O)
                        If O <> P Then
                            K += 1
                            G.Positions(I) = P.ToVector3
                        End If
                    Next
                End If
            Next
        Else
            For Each node In SceneModel.SceneNode.Traverse(False)
                If InStr(LCase(node.GetType.Name), "meshnode") > 0 Then
                    N = node
                    M = N.ModelMatrix.ToMatrix3D
                    T.Matrix = M
                    G = N.Geometry
                    J = G.Positions.Count - 1
                    For I = 0 To J
                        O = G.Positions(I).ToPoint3D
                        P = T.Transform(O)
                        If O <> P Then
                            K += 1
                            G.Positions(I) = P.ToVector3
                        End If
                    Next
                End If
            Next
        End If
        If K = 0 Then MsgBox("geen punten hervormt")
    End Sub

    Sub ModelPartTurn()

        ' draait het volgende onderdeel

        ' W(0) = "turn"
        ' V(1) = X centrum
        ' V(2) = Y centrum
        ' V(3) = Z centrum
        ' V(4) = van hoek
        ' V(5) = tot hoek
        ' V(6) = om xas = 1
        ' V(7) = om yas = 1
        ' V(8) = om zas = 1
        ' V(9) = secondes
        ' V(10) = Type, 0, 1, 2
        ' V(11) = bewaar

        For I = 0 To FI
            V(1) += OX(I)
            V(2) += OY(I)
            V(3) += OZ(I)
        Next
        PartTurn.CX = V(1)
        PartTurn.CY = V(2)
        PartTurn.CZ = V(3)
        PartTurn.B = V(4)
        PartTurn.E = V(5)
        PartTurn.AX = V(6)
        PartTurn.AY = V(7)
        PartTurn.AZ = V(8)
        PartTurn.T = V(9)
        PartTurn.R = V(10)
        If V(11) = 0 Then
            PartTransform.Children.Clear()
        End If
    End Sub

    Sub ModelPartWall()

        ' maakt een muur

        ' W(0) = "wall"
        ' V(1) = maat
        ' V(2) = aantal vlakken

        Dim B As New MeshBuilder
        Dim N As Short = V(2)
        Dim T As New Transform3DGroup
        Dim H As Single = V(1) / N
        Dim F As Single = 1 / N

        ModelPartNew()

        If W(6) = "" Then
            For X = 0 To N - 1
                For Y = 0 To N - 1
                    B.AddQuad(New Vector3(X * H, (Y + 1) * H, 0),
                          New Vector3(X * H, Y * H, 0),
                          New Vector3((X + 1) * H, Y * H, 0),
                          New Vector3((X + 1) * H, (Y + 1) * H, 0),
                          New Vector2(X * F, (N - 1 - Y) * F),
                          New Vector2(X * F, (N - 1 - Y + 1) * F),
                          New Vector2((X + 1) * F, (N - 1 - Y + 1) * F),
                          New Vector2((X + 1) * F, (N - 1 - Y) * F))

                Next
            Next
        Else
            For X = 0 To N - 1
                For Y = 0 To N - 1
                    B.AddQuad(New Vector3(X * H, 0, -(Y + 1) * H),
                          New Vector3(X * H, 0, -Y * H),
                          New Vector3((X + 1) * H, 0, -Y * H),
                          New Vector3((X + 1) * H, 0, -(Y + 1) * H),
                          New Vector2(X * F, (N - 1 - Y) * F),
                          New Vector2(X * F, (N - 1 - Y + 1) * F),
                          New Vector2((X + 1) * F, (N - 1 - Y + 1) * F),
                          New Vector2((X + 1) * F, (N - 1 - Y) * F))

                Next
            Next
        End If


        ModelPart.Geometry = B.ToMeshGeometry3D()
        ModelPart.Material = PartMaterial


        FI += 1
        OX(FI) = V(3)
        OY(FI) = V(4)
        OZ(FI) = V(5)
        AX(FI) = 0 'V(6)
        AY(FI) = V(7)
        AZ(FI) = V(8)
        If W(9) <> "" Then SX(FI) = V(9)
        If W(10) <> "" Then SY(FI) = V(10)
        If W(11) <> "" Then SZ(FI) = V(11)
        For I = FI To 0 Step -1
            ModelPart.Translate(T, OX(I), OY(I), OZ(I))
            ModelPart.Rotate(T, New Point3D(OX(I), OY(I), OZ(I)), AX(I), AY(I), AZ(I))
            ModelPart.Scale(T, New Vector3D(SX(I), SY(I), SZ(I)))
        Next
        FI -= 1
        ModelPart.Transform = T
        ModelPart.Add()
    End Sub

    Function PartMoveGet() As Transform3DGroup
        '
        ' animeert een verplaatsing van een onderdeel
        '
        Dim T As New Transform3DGroup

        If PartMove.T > 0 Then
            Dim TX As New TranslateTransform3D
            Dim TY As New TranslateTransform3D
            Dim TZ As New TranslateTransform3D

            Dim AX As New DoubleAnimation(PartMove.X, 0, New Windows.Duration(TimeSpan.FromSeconds(PartMove.T)))
            Dim AY As New DoubleAnimation(PartMove.Y, 0, New Windows.Duration(TimeSpan.FromSeconds(PartMove.T)))
            Dim AZ As New DoubleAnimation(PartMove.Z, 0, New Windows.Duration(TimeSpan.FromSeconds(PartMove.T)))

            If PartMove.R > 0 Then
                AX.RepeatBehavior = RepeatBehavior.Forever
                AY.RepeatBehavior = RepeatBehavior.Forever
                AZ.RepeatBehavior = RepeatBehavior.Forever
                If PartMove.R > 1 Then
                    AX.AutoReverse = True
                    AY.AutoReverse = True
                    AZ.AutoReverse = True
                End If
            End If
            TX.BeginAnimation(TranslateTransform3D.OffsetXProperty, AX)
            TY.BeginAnimation(TranslateTransform3D.OffsetYProperty, AY)
            TZ.BeginAnimation(TranslateTransform3D.OffsetZProperty, AZ)
            T.Children.Add(TX)
            T.Children.Add(TY)
            T.Children.Add(TZ)
        End If
        PartMoveGet = T
    End Function
    '
    Function PartTurnGet() As Transform3DGroup
        '
        ' animeert een verdraaiing van een onderdeel
        '
        Dim T As New Transform3DGroup

        If PartTurn.T > 0 Then ' als de tijd groter dan 0 is

            Dim A As New DoubleAnimation
            Dim R As New RotateTransform3D
            '
            A.From = PartTurn.B
            A.To = PartTurn.E
            A.Duration = TimeSpan.FromSeconds(PartTurn.T)
            If PartTurn.R > 0 Then
                A.RepeatBehavior = RepeatBehavior.Forever
                If PartTurn.R > 1 Then
                    A.AutoReverse = True
                End If
            End If

            R.CenterX = PartTurn.CX
            R.CenterY = PartTurn.CY
            R.CenterZ = PartTurn.CZ
            R.Rotation = New AxisAngleRotation3D(New Vector3D(PartTurn.AX, PartTurn.AY, PartTurn.AZ), 0)
            R.Rotation.BeginAnimation(AxisAngleRotation3D.AngleProperty, A)
            T.Children.Add(R)
            If PartTransform.Children.Count > 0 Then
                T.Children.Add(PartTransform)
            Else
                PartTransform.Children.Add(T)
            End If
        End If
        Return T
    End Function

    Public Structure TMove

        ' type voor verplaatsingen

        Dim T As Single ' tijd
        Dim R As Short ' herhaal gedrag
        Dim X As Single
        Dim Y As Single
        Dim Z As Single
    End Structure

    Public Structure TTurn

        ' type voor verdraaiing

        Dim T As Single ' tijd in secondes
        Dim R As Short ' herhaal gedrag
        Dim CX As Single ' X centrum
        Dim CY As Single
        Dim CZ As Single
        Dim AX As Single ' X as
        Dim AY As Single
        Dim AZ As Single
        Dim B As Single ' beginhoek
        Dim E As Single ' eindhoek
    End Structure

    Sub PartLegoInit()

        ' start lego onderdeel

        PartBuilder = New MeshBuilder

        For I = 0 To 99
            LegoParts(I).Name = ""
        Next
        StepHeight = 0
        PartData = False
    End Sub

    Sub PartLegoDone()

        ' stopt lego onderdeel

        If PartBuilder IsNot Nothing Then
            If PartBuilder.ToMeshGeometry3D.Positions.Count > 0 Then
                ModelPartNew()
                ModelPart.Geometry = PartBuilder.ToMeshGeometry3D()
                ModelPart.Material = PartMaterial
                ModelPart.Add()
            End If
        End If
    End Sub

    Sub PartLegoAdd()

        ' voegt lego onderdeel toe aan lijst

        Dim I As Short
        Dim S As String

        S = Mid(ParsFile(FI - 1), FileRoot.SelectedItem.Length + 4)
        S = Left(S, S.Length - 4)
        For I = 0 To 99
            If LegoParts(I).Name = S Then ' als onderdeel al in de lijst staat
                LegoParts(I).Numb += 1
                Exit Sub
            End If
        Next
        I = 0
        While LegoParts(I).Name <> ""
            I += 1
        End While
        LegoParts(I).Name = S
        LegoParts(I).Numb = 1
    End Sub

    Sub PartLegoComment()

        ' bepaalt opdrachten in lego commentaar

        ' W(0) = "0"

        If W(1) = "BFC" Then
            If W(2) = "INVERTNEXT" Then
                INV(MI) = True
            ElseIf W(2) = "CERTIFY" Then
                If W(3) = "CW" Then
                    If INV(MI) Then
                        REV(MI) = False
                    Else
                        REV(MI) = True
                    End If
                ElseIf W(3) = "CCW" Then
                    If INV(MI) Then
                        REV(MI) = True
                    Else
                        REV(MI) = False
                    End If
                End If
            End If
        End If
    End Sub

    Sub PartLegoFile()

        ' bepaalt opdrachten in lego bestand

        ' W(0) = "1"
        ' V(1) = kleur
        ' V(2) .. V(13) = matrix
        ' W(14) = bestand.dat

        MI += 1 ' als dit 1e file is dan matrix = 1

        ML(MI).OffsetX = V(2)
        ML(MI).OffsetY = V(3)
        ML(MI).OffsetZ = V(4)
        ML(MI).M11 = V(5)
        ML(MI).M21 = V(6)
        ML(MI).M31 = V(7)
        ML(MI).M12 = V(8)
        ML(MI).M22 = V(9)
        ML(MI).M32 = V(10)
        ML(MI).M13 = V(11)
        ML(MI).M23 = V(12)
        ML(MI).M33 = V(13)
        ML(MI).M14 = 0
        ML(MI).M24 = 0
        ML(MI).M34 = 0
        ML(MI).M44 = 1

        INV(MI) = INV(MI - 1) ' als file = 2 krijgt inv van file 1,  file 3 van file 2 file 3 
        If W(14) <> "" Then
            ParsFile(FI) = "D:\Data\ldraw\parts\" & W(14)
            If File.Exists(ParsFile(FI)) = False Then
                ParsFile(FI) = "D:\Data\ldraw\p\" & W(14)
                If File.Exists(ParsFile(FI)) = False Then
                    MsgBox(ParsFile(FI) & " niet gevonden")
                End If
            End If
            If ParsSkip = False Then
                TextParsRuns(ParsFile(FI))
            Else
                Exit Sub
            End If
        End If
        If MI > 1 Then
            MI -= 1
            INV(MI) = INV(MI - 1)
        End If
    End Sub

    Sub PartLegoTriangle()

        ' maakt een driehoek

        ' W(0) = "3"
        ' V(1) = kleur

        Dim I As Short
        Dim D As Single

        For I = 1 To 3
            PA(I).X = V(I * 3 - 1)
            PA(I).Y = V(I * 3 + 0)
            PA(I).Z = V(I * 3 + 1)
            PartLegoCalc(I)
        Next

        D = ML(0).Determinant
        For I = 1 To MI
            D *= ML(I).Determinant
        Next

        If Not REV(MI) Then
            If D > 0 Then
                PartLegoFace(1, 2, 3, 0)
            Else
                PartLegoFace(3, 2, 1, 0)
            End If
        Else ' als omgekeerd
            If D > 0 Then
                PartLegoFace(3, 2, 1, 0)
            Else
                PartLegoFace(1, 2, 3, 0)
            End If
        End If
    End Sub

    Sub PartLegoRectangle()

        ' maakt een rechthoek

        ' V(0) = 4
        ' V(1) = kleur
        ' V(2) = X1
        ' V(3) = Y1
        ' V(4) = Z1

        Dim I As Short
        Dim D As Single

        For I = 1 To 4
            PA(I).X = V(I * 3 - 1) ' V(2, 5,  8, 11)
            PA(I).Y = V(I * 3 + 0) ' V(3, 6,  9, 12)
            PA(I).Z = V(I * 3 + 1) ' V(4, 7, 10, 13)
            PartLegoCalc(I)
        Next

        D = ML(0).Determinant
        For I = 1 To MI
            D *= ML(I).Determinant
        Next

        If Not REV(MI) Then
            If D > 0 Then
                PartLegoFace(1, 2, 3, 4)
            Else
                PartLegoFace(4, 3, 2, 1)
            End If
        Else
            If D > 0 Then
                PartLegoFace(4, 3, 2, 1)
            Else
                PartLegoFace(1, 2, 3, 4)
            End If
        End If
    End Sub

    Sub PartLegoCalc(N As Long)

        ' berekent lego afmetingen

        Dim P As New Media3D.Point3D

        P.X = PA(N).X
        P.Y = PA(N).Y
        P.Z = PA(N).Z

        For I = MI To 1 Step -1
            P = Media3D.Point3D.Multiply(P, ML(I))
        Next

        PA(N).X = OX(0) + PartLegoSize(P.X)
        PA(N).Y = OY(0) + PartLegoSize(P.Y)
        PA(N).Z = OZ(0) + PartLegoSize(P.Z)
    End Sub

    Sub ModelExportBin()

        ' maakt een .bin bestand van een onderdeel

        Dim F As New BinaryWriter(File.Open(FileRoot.SelectedItem & "\" & GetFilePath(ParsFile(FI)) & PartFile & ".bin", FileMode.Create))
        Dim I As Long
        Dim J As Long
        Dim K As Long
        Dim M As HelixToolkit.Wpf.SharpDX.GeometryModel3D
        Dim N As Long
        Dim O As Long
        Dim P As HelixToolkit.Wpf.SharpDX.MeshGeometry3D

        K = SceneModel.SceneNode.Items.Count ' aantal onderdelen

        For J = 0 To K - 1 ' voor alle onderdelen
            M = SceneModel.SceneNode.Items(J) ' onderdeel
            P = M.Geometry ' geometrie
            N = P.TriangleIndices.Count ' aantal driehoek punten
            For I = 0 To N - 1 ' voor alle driehoek punten
                O = P.TriangleIndices(I) ' werkelijke punt
                F.Write(P.Positions(O).X)
                F.Write(P.Positions(O).Y)
                F.Write(P.Positions(O).Z)
            Next
        Next
        F.Close()
    End Sub

    Sub PartLegoFace(P1 As Short, P2 As Short, P3 As Short, P4 As Short)

        ' voegt vlak toe aan lego onderdeel

        If P4 = 0 Then
            PartBuilder.AddTriangle(PA(P1), PA(P2), PA(P3))
        Else
            PartBuilder.AddQuad(PA(P1), PA(P2), PA(P3), PA(P4))
        End If
    End Sub

    Sub PartLegoBin(S As String)

        ' tekent een lego onderdeel van een bin bestand

        Dim I1 As Long
        Dim I2 As Long
        Dim I3 As Long

        If File.Exists(S) Then

            Dim B As New HelixToolkit.Wpf.SharpDX.MeshBuilder
            Dim F As New FileInfo(S)
            Dim I As Long
            'Dim L As Long
            Dim N As Long
            Dim T As New Media3D.Transform3DGroup

            ModelPartNew()

            Using R As New BinaryReader(File.Open(S, FileMode.Open))

                N = F.Length / 12 ' aantal bytes -> aantal punten een punt bestaat uit 3 singles van ieder 4 bytes = 12 bytes
                For I = 1 To N
                    PA(I) = New SharpDX.Vector3(R.ReadSingle(), R.ReadSingle(), R.ReadSingle()) * SC
                Next
            End Using

            N /= 3 ' aantal driehoeken
            For I = 1 To N
                I1 = (I - 1) * 3 + 1
                I2 = (I - 1) * 3 + 2
                I3 = (I - 1) * 3 + 3
                B.AddTriangle(PA(I1), PA(I2), PA(I3))
            Next

            ModelPart.Geometry = B.ToMeshGeometry3D()
            ModelPart.Material = PartMaterial
            For I = FI To 0 Step -1
                ModelPart.Translate(T, OX(I), OY(I), OZ(I))
                ModelPart.Rotate(T, New Media3D.Point3D(OX(I), OY(I), OZ(I)), AX(I), AY(I), AZ(I))
            Next
            ModelPart.Transform = T
            ModelPart.Add()
        Else
            MsgBox(S & " niet gevonden " & vbCrLf & ParsFile(FI - 1))
            Interaction.Shell("C:\windows\system32\notepad.exe " & """" & ParsFile(FI - 1) & """", vbNormalFocus)
            ParsSkip = True
        End If
    End Sub

    Function PartLegoSize(R As Single) As Single

        ' verkleint afmetingen

        If R > 0.1 Then
            PartLegoSize = (R - 0.1) / 1000
        Else
            PartLegoSize = (R + 0.1) / 1000
        End If
    End Function

    Sub PartLegoStep()

        ' volgende stap

        ' W(0) = "step"
        ' V(1) = index of hoogte
        ' W(2) = [wacht]

        If W(2) = "" Then
            StepHeight = V(1)
        Else
            StepHeight = V(2)
        End If
        If W(2) <> "" Then Wait(V(2))
    End Sub

    Sub PartDataList()

        ' maakt een lijst met de lego onderdelen van een model

        Dim D As New DataView
        Dim I As Short
        Dim J As Short
        Dim N As Short = 0
        Dim F As String
        Dim P As String
        Dim R As DataRow
        Dim S As String = ""
        Dim T() As String

        If LegoParts(0).Name <> "" Then
            D.Table = New DataTable("sort")
            D.Table.Columns.Add("numb")
            D.Table.Columns.Add("name")
            I = 0
            While LegoParts(I).Name <> ""
                P = LegoParts(I).Name
                P = Left(P, 1) & "\" & P
                P = GetFilePath(P)
                P = FileRoot.SelectedItem & "\" & P & "#\info.txt"
                F = GetFileName(LegoParts(I).Name)
                If File.Exists(P) Then
                    T = File.ReadAllLines(P)
                    J = UBound(T)
                    If J > -1 Then
                        N += LegoParts(I).Numb
                        R = D.Table.NewRow
                        R(0) = LegoParts(I).Numb
                        R(1) = F & " " & T(0)
                        D.Table.Rows.Add(R)
                    End If
                Else
                    MsgBox("Onderdeel " & P & " heeft nog geen naam")
                    ParsSkip = True
                    Exit Sub
                End If
                I += 1
            End While
            D.Sort = "name ASC"
            For Each row In D
                S &= LE(row(0)) & row(1) & vbCrLf
            Next
            S &= vbCrLf
            S &= LE(N) & "totaal" & vbCrLf
            File.WriteAllText(FileRoot.SelectedItem & "\" & GetFilePath(ParsFile(FI)) & "data.txt", S)
            PartData = False
        End If
    End Sub

    Structure TLegoParts
        Dim Name As String
        Dim Numb As Short
    End Structure

    Public INV(20) As Boolean '
    Public REV(20) As Boolean ' counter clock wise; default false
    Public ML(20) As Matrix3D ' matrix array
    Public MI As Short ' matrix index
    Public PartBuilder As MeshBuilder
    Public StepHeight As Single
    Public LegoParts(100) As TLegoParts
    Public PartData As Boolean
    Public PartFile As String

    Public PA(200000) As SharpDX.Vector3 ' punt array
    Public UA(200000) As SharpDX.Vector2 ' uv array
    Public NA(200000) As SharpDX.Vector3 ' normaal array
    Public PN As Long ' point index
    Public UN As Long ' uvmap index
    Public NN As Long ' normal index
    Public PMax As Vector3
    Public PMin As Vector3
    Public PartTransform As New Transform3DGroup
    Public PartGeo As New HelixToolkit.Wpf.SharpDX.MeshGeometry3D

    Public PartMove As TMove
    Public PartTurn As TTurn
    Public PartKeep As Boolean
End Module
