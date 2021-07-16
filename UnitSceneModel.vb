Imports System.IO
Imports System.Windows.Controls
Imports System.Windows.Media.Media3D
Imports HelixToolkit.Wpf.SharpDX
Imports HelixToolkit.Wpf.SharpDX.Animations
Imports HelixToolkit.Wpf.SharpDX.Assimp
Imports HelixToolkit.Wpf.SharpDX.Model.Scene
Imports SharpDX

Module UnitSceneModel

    Public SceneModel As New SceneNodeGroupModel3D()

    Class ClassSceneModel
        ' Inherits SceneNodeGroupModel3D ' niet mogelijk

    End Class

    Sub SceneModelInit()

        ' start de objecten in een ruimte

        Dim K As Keyframe

        ModelPart.Init()

        WalkFile = ""
        IdleFile = ""
        ModelTrans.Children.Clear()
        ModelTrans.Children.Add(New RotateTransform3D)
        ModelTrans.Children.Add(New TranslateTransform3D)
        ModelTrans.Children.Add(New TranslateTransform3D)
        ModelTrans.Children.Add(New TranslateTransform3D)
        ModelTrans.Children.Add(New ScaleTransform3D)
        Sync.B = -1 ' synchroniseer bot
        For I = 0 To 4
            AnimUpdt(I) = Nothing ' clear old animation
            AE(I) = 0 ' stop animaties
            AC(I) = 0
            For J = 1 To 10
                FB(I, J).I = 0
                FB(I, J).B = 400
            Next
            BoneIndex(I) = 0
            For J = 0 To 400
                Bone(I, J).I = J
            Next
        Next
        SC = 1
        GS = 1
        For I = 0 To 20
            OX(I) = 0
            OY(I) = 0
            OZ(I) = 0
            AX(I) = 0
            AY(I) = 0
            AZ(I) = 0
            SX(I) = 1
            SY(I) = 1
            SZ(I) = 1
        Next
        SE = False
        AI = 0
        SI = 0
        BoneTarget = 0
        Navigate = False
        MA = 0 ' Object angle y
        MY = 0
        MZ = 0
        MP = New Vector3D(0, 0, 0)
        K.Rotation = New SharpDX.Quaternion(0, 0, 0, 1)
        K.Translation = New Vector3(0, 0, 0)
        K.Scale = New Vector3(1, 1, 1)
        For I = 0 To 400
            For J = 0 To 400
                KF(I, J) = K
            Next
        Next
        Paint = False
        Up = "Y"
    End Sub
    '
    Sub SceneModelExport()

        ' exporteert een model of onderdelen

        ' W(0) = "export"
        ' W(1) = filename
        ' W(2) = [objecten]
        ' V(3) = [schaal}

        Dim I As Short
        Dim J As Short
        Dim N As Short
        Dim S As String = ""
        Dim T As String

        PartLegoDone()
        W(1) = GetFileAddress(W(1)) ' file
        T = GetFileType(W(1))
        Select Case T
            Case ".bin"
                PartFile = GetFileName(W(1))
                ModelExportBin()
                Exit Sub
            Case ".dae"
                T = "collada"
            Case ".obj"
                T = "obj"
            Case ".objnomtl"
                T = "objnomtl"
                W(1) = Left(W(1), Len(W(1)) - 5)
        End Select
        If W(2) = "" Then ' als een model niet gespiltst hoeft te worden
            Using E As New HelixToolkit.Wpf.SharpDX.Assimp.Exporter
                If ModelAnim(SI) IsNot Nothing Then ' als het model is
                    S = E.ExportToFile(W(1), ModelAnim(SI), T).ToString
                Else ' exporteer de onderdelen
                    N = SceneModel.SceneNode.ItemsCount
                    S = E.ExportToFile(W(1), SceneModel.SceneNode, T).ToString ' moeten eerst getransformeerd worden
                End If
            End Using
        ElseIf W(2) = V(2).ToString Then ' splits model met nummers
            J = 0
            Using E As New HelixToolkit.Wpf.SharpDX.Assimp.Exporter
                For Each Node In ModelAnim(SI).Root.Traverse(True) ' aantal objecten
                    J += 1
                    If J = V(2) Then
                        W(1) = FileRoot.SelectedItem & GetFilePath(W(1)) & J & GetFileType(W(1))
                        S = E.ExportToFile(W(1), Node, T).ToString
                    End If
                Next
            End Using
        Else ' slits model met namen
            Using E As New HelixToolkit.Wpf.SharpDX.Assimp.Exporter
                N = ModelAnim(SI).Root.ItemsCount - 1 ' aantal objecten
                For I = 0 To N
                    S = Strings.Replace(W(1), ".", ModelAnim(SI).Root.Items(I).Name & ".") ' voeg de objectnaam toe aan de filenaam
                    S = E.ExportToFile(S, ModelAnim(SI).Root.Items(I), T).ToString
                Next
            End Using
        End If
        If S <> "Succeed" Then MsgBox(S)
    End Sub

    Sub SceneModelImport()

        ' importeert een model uit een .dae, .fbx of .obj bestand

        ' W(0) = "import"
        ' W(1) = filename
        ' V(2) = [dx]
        ' V(3) = [dy]
        ' V(4) = [dz]
        ' V(5) = [hx]
        ' V(6) = [hy]
        ' V(7) = [hz]
        ' V(8) = [sx]
        ' V(9) = [sy]
        ' V(10) = [sz]

        Dim E As String ' error
        Dim N As MeshNode
        Dim S As HelixToolkitScene
        Dim T As New Transform3DGroup

        W(1) = GetFileAddress(W(1))
        If File.Exists(W(1)) Then
            If GetFileType(W(1)) = ".obj" Then
                Using ObjectRead As New Importer
                    S = ObjectRead.Load(W(1))
                End Using
                For Each node In S.Root.Traverse(True)
                    If InStr(LCase(node.GetType.ToString), "meshnode") > 0 Then
                        N = node
                        N.Material = PartMaterial
                    End If
                Next
                FI += 1
                OX(FI) = V(2)
                OY(FI) = V(3)
                OZ(FI) = V(4)
                AX(FI) = V(5)
                AY(FI) = V(6)
                AZ(FI) = V(7)
                If W(8) <> "" Then SX(FI) = V(8) Else SX(FI) = 1
                If W(9) <> "" Then SY(FI) = V(9) Else SY(FI) = 1
                If W(10) <> "" Then SZ(FI) = V(10) Else SZ(FI) = 1
                For I = FI To 0 Step -1
                    ModelPart.Translate(T, OX(I), OY(I), OZ(I))
                    ModelPart.Rotate(T, New Point3D(OX(I), OY(I), OZ(I)), AX(I), AY(I), AZ(I))
                    ModelPart.Scale(T, New Vector3D(SX(I), SY(I), SZ(I)))
                Next
                FI -= 1
                S.Root.ModelMatrix = T.ToMatrix
                SceneModel.AddNode(S.Root)
            Else ' voor .dae en .fbx bestanden
                SI += 1
                Using ObjectRead As New Importer
                    E = ObjectRead.Load(W(1), ModelAnim(SI)).ToString
                End Using
                If E = "Succeed" Then
                    FI += 1
                    OX(FI) = V(2)
                    OY(FI) = V(3)
                    OZ(FI) = V(4)
                    AX(FI) = V(5)
                    AY(FI) = V(6)
                    AZ(FI) = V(7)
                    If W(8) <> "" Then SX(FI) = V(8) Else SX(FI) = 1
                    If W(9) <> "" Then SY(FI) = V(9) Else SY(FI) = 1
                    If W(10) <> "" Then SZ(FI) = V(10) Else SZ(FI) = 1
                    For I = FI To 0 Step -1
                        ModelPart.Translate(T, OX(I), OY(I), OZ(I))
                        ModelPart.Rotate(T, New Point3D(OX(I), OY(I), OZ(I)), AX(I), AY(I), AZ(I))
                        ModelPart.Scale(T, New Vector3D(SX(I), SY(I), SZ(I)))
                    Next
                    FI -= 1
                    ModelAnim(SI).Root.ModelMatrix = T.ToMatrix
                    SceneModel.AddNode(ModelAnim(SI).Root) ' voeg het object toe aan de scene
                    If MediaScene.IsShadowMappingEnabled Then
                        For Each Node In ModelAnim(SI).Root.Items.Traverse(False)
                            If Node.GetType = GetType(BoneSkinMeshNode) Then
                                N = Node
                                N.IsThrowingShadow = True
                            End If
                        Next
                    End If
                Else
                    MsgBox(E)
                    ParsSkip = True
                End If
            End If
        Else
            MsgBox("bestand " & W(1) & " niet gevonden")
            ParsSkip = True
        End If
    End Sub

    Sub SceneModelMatrix()

        ' importeert een matrix 

        ' W(0) = "matrix"
        ' V(1) = index
        ' V(2)..V(17) = waardes

        Dim K As Keyframe
        Dim M As SharpDX.Matrix
        Dim N As Short
        Dim T As Vector3

        If W(5) = "" Then
            T = New Vector3(V(2), V(3), V(4))
            N = AnimUpdt(AI).Animation.NodeAnimationCollection(V(1)).KeyFrames.Count
            For I = 0 To N - 1
                K = AnimUpdt(AI).Animation.NodeAnimationCollection(V(1)).KeyFrames(I)
                K.Translation += T
                AnimUpdt(AI).Animation.NodeAnimationCollection(V(1)).KeyFrames(I) = K
            Next
        Else
            M.M11 = V(2)
            M.M21 = V(3)
            M.M31 = V(4)
            M.M41 = V(5)
            M.M12 = V(6)
            M.M22 = V(7)
            M.M32 = V(8)
            M.M42 = V(9)
            M.M13 = V(10)
            M.M23 = V(11)
            M.M33 = V(12)
            M.M43 = V(13)
            M.M14 = V(14)
            M.M24 = V(15)
            M.M34 = V(16)
            M.M44 = V(17)
            AnimUpdt(AI).Animation.NodeAnimationCollection(Bone(AI, V(1)).I).Node.ModelMatrix = M
        End If
    End Sub

    Sub SceneModelMatrices()

        ' exporteert de modelmatrix van ieder bot

        ' W(0) = "matrices'
        ' W(1) = bestandnaam
        ' V(2) = [type]
        ' V(3) = [verschil]
        ' V(4) = [nieuw]

        Dim M As SharpDX.Matrix
        Dim N As Short

        W(1) = GetFileAddress(W(1))
        FileOpen(1, W(1), OpenMode.Output)
        N = AnimUpdt(AI).Animation.NodeAnimationCollection.Count ' aantal botten
        'If W(4) <> "" Then
        For I = 0 To N - 1
            If V(2) = 0 Then
                MM(AI, I) = AnimUpdt(AI).Animation.NodeAnimationCollection(I).Node.ModelMatrix
                If V(3) <> 0 Then MM(V(3), I) = AnimUpdt(V(3)).Animation.NodeAnimationCollection(I).Node.ModelMatrix
            Else
                MM(AI, I) = AnimUpdt(AI).Animation.NodeAnimationCollection(I).Node.TotalModelMatrix
                If V(3) <> 0 Then MM(V(3), I) = AnimUpdt(V(3)).Animation.NodeAnimationCollection(I).Node.TotalModelMatrix
            End If
        Next
        'End If
        For I = 0 To N - 1
            PrintLine(1, "' " & AnimUpdt(AI).Animation.NodeAnimationCollection(I).Node.Name)
            If V(3) = 0 Then
                M = MM(AI, I)
                PrintLine(1, "matrix" & LB(I) &
                      SF(M.M11) & SF(M.M21) & SF(M.M31) & SF(M.M41) &
                      SF(M.M12) & SF(M.M22) & SF(M.M32) & SF(M.M42) &
                      SF(M.M13) & SF(M.M23) & SF(M.M33) & SF(M.M43) &
                      SF(M.M14) & SF(M.M24) & SF(M.M34) & SF(M.M44))
            Else
                M = MM(AI, I) - MM(V(2), I)
                PrintLine(1, "matrix" & LB(I) & SF(M.M41) & SF(M.M42) & SF(M.M43))
            End If

        Next
        FileClose(1)
    End Sub

    Sub SceneModelNodes()

        ' exporteert de nodes van een model

        ' W(0) = "nodes"
        ' W(1) = bestandnaam
        ' V(2) = type, 0 = compleet, 1 = alleen translatie vector

        Dim I As Short = 0
        Dim T As Vector3

        W(1) = GetFileAddress(W(1))
        FileOpen(1, W(1), OpenMode.Output)
        For Each Node In ModelAnim(SI).Root.Traverse(True)
            I += 1
            T = Node.ModelMatrix.TranslationVector
            If V(2) = 0 Then
                PrintLine(1, I & ", " & Node.GetType.ToString & ", " & Node.Name & ", " & Node.Parent.Name & ", " & T.ToString)
            Else
                PrintLine(1, I & SF(T.X) & SF(T.Y) & SF(T.Z))
            End If
        Next
        FileClose(1)
    End Sub

    Sub SceneModelShow()

        '

        ModelPos.Text = "model " & SF(MP.X) & SF(MP.Y) & SF(MP.Z) & " modelhoek " & MA & " camerahoek " & SceneCamera.CA
    End Sub

    Sub SceneModelWires()

        ' toont het draadmodel van alle objecten

        ' W(0) = "wires"

        Dim M As MeshNode

        For Each Node In SceneModel.SceneNode.Traverse(True)
            If InStr(LCase(Node.GetType.Name), "meshnode") > 0 Then
                M = Node
                M.RenderWireframe = True
                M.IsHitTestVisible = True
            End If
        Next
    End Sub

    Public MR(4) As SharpDX.Matrix
    Public ModelPos As New TextBox
    Public ModelLabel As New Label
End Module
