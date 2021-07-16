Imports HelixToolkit.Wpf.SharpDX
Imports HelixToolkit.Wpf.SharpDX.Animations
Imports HelixToolkit.Wpf.SharpDX.Assimp
Imports HelixToolkit.Wpf.SharpDX.Controls
Imports SharpDX
Imports System.Diagnostics
Imports System.IO
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Windows.Media.Media3D

Module UnitModelPose


    Public ModelAnim(4) As HelixToolkitScene

    Sub ModelAnimStart()

        ' start een animatie

        ' W(0) = "animate"
        ' V(1) = index
        ' V(2) = [orientatie 1 of -1]
        ' V(3) = [reset keyframes]
        ' V(4) = [eind tijd]

        Dim I As Short
        Dim J As Short
        Dim K As Keyframe
        Dim M As Short
        Dim N As Short      ' aantal botten

        If V(1) < 0 Then ' Stop alle animaties
            For I = 0 To 4
                AE(I) = 0
            Next
            Exit Sub
        End If

        If ModelAnim(SI) IsNot Nothing Then
            If ModelAnim(SI).HasAnimation Then
                MR(SI) = ModelAnim(SI).Root.ModelMatrix
                AI += 1
                N = ModelAnim(SI).Animations.Count
                If V(1) > N - 1 Then V(1) = N - 1
                AnimUpdt(AI) = New NodeAnimationUpdater(ModelAnim(SI).Animations(V(1)))
            End If
            If AnimUpdt(AI) IsNot Nothing Then
                If W(2) = "" Then
                    AC(AI) = 1
                Else
                    AC(AI) = V(2)
                End If
                N = AnimUpdt(AI).Animation.NodeAnimationCollection.Count ' aantal botten
                For I = 0 To N - 1
                    Bone(AI, I).I = I ' onthoud de botten
                    Bone(AI, I).O = AnimUpdt(AI).Animation.NodeAnimationCollection(I).Node.Name
                    Bone(AI, I).N = Bone(AI, I).O
                    MM(AI, I) = AnimUpdt(AI).Animation.NodeAnimationCollection(I).Node.ModelMatrix ' onthoudt de modelmatrix van ieder bot
                Next

                If W(3) <> "" Then ' reset keyframes
                    If V(4) = 0 Then M = 5 Else M = V(4) ' aantal keyframes
                    If W(5) = "" Then V(5) = 2 ' eindtijd
                    AnimUpdt(AI).Animation.StartTime = 0
                    AnimUpdt(AI).Animation.EndTime = V(5)
                    For I = 0 To N - 1 ' voor ieder bot
                        AnimUpdt(AI).Animation.NodeAnimationCollection(I).KeyFrames.Clear() ' wis alle keyframes
                        K.BoneIndex = I
                        K.Scale = New SharpDX.Vector3(1, 1, 1)
                        K.Translation = MM(AI, I).TranslationVector ' houdt verplaatsing
                        If V(3) = 0 Then ' wis rotaties
                            K.Rotation = New SharpDX.Quaternion(0, 0, 0, AC(AI))
                        Else ' houdt rotatie
                            K.Rotation = MatrixToQuaternion(MM(AI, I))
                            K.Rotation.W *= AC(AI)
                        End If
                        For J = 0 To M - 1 ' maak nieuwe keyframes
                            K.Time = V(5) * J / (M - 1)
                            AnimUpdt(AI).Animation.NodeAnimationCollection(I).KeyFrames.Add(K)
                        Next
                    Next
                End If
                For I = 0 To 4 ' start animaties
                    AE(I) = True
                Next
            Else
                MsgBox("model kan niet bewegen")
            End If
        Else
            MsgBox("geen model geladen")
        End If
    End Sub

    Sub ModelAnimPose()

        ' stelt de houding van een bot in

        ' W(0) = "pose"
        ' V(1) = model nummer
        ' V(2) = bot nummer
        ' V(3) = type houding / beweging
        ' V(4) = dx
        ' V(5)
        ' V(6)
        ' V(7) = ax
        ' V(8)
        ' V(9)
        ' V(10) = fdx  factor
        ' V(11) = fdy
        ' V(12) = fdz
        ' V(13) = fax
        ' V(14) = fay
        ' V(15) = faz
        ' V(16) = eerste keyframe
        ' V(17) = laatste keyframe

        Dim I As Short
        Dim J As Short
        Dim N As Short

        If AI > 0 Then
            Pose.A = V(1)
            If AnimUpdt(Pose.A) IsNot Nothing Then
                N = AnimUpdt(Pose.A).Animation.NodeAnimationCollection.Count ' aantal botten
                If W(2) = V(2).ToString Then ' als een bot nummer opgegeven is
                    Pose.B = Bone(Pose.A, V(2)).I
                Else ' als een bot naam opgegeven is
                    W(2) = GetOldBone(W(2))
                    For I = 0 To N
                        If LCase(AnimUpdt(Pose.A).Animation.NodeAnimationCollection(I).Node.Name) = LCase(W(2)) Then
                            Pose.B = I
                            GoTo verder
                        End If
                    Next
                    MsgBox("bot " & W(2) & " niet gevonden")
                    ParsSkip = True
                    Exit Sub
                End If
verder:
                Pose.T = V(3)
                SE = False
                SDX.Minimum = -1 / SC
                SDX.Maximum = 1 / SC
                SDY.Minimum = -1 / SC
                SDY.Maximum = 1 / SC
                SDZ.Minimum = -1 / SC
                SDZ.Maximum = 1 / SC

                SDX.Value = V(4) / SC
                SDY.Value = V(5) / SC
                SDZ.Value = V(6) / SC
                If Up = "Y" Then
                    SAX.Value = V(7)
                    SAY.Value = V(8)
                    SAZ.Value = V(9)
                Else
                    SAX.Value = V(7) * UpX
                    SAY.Value = V(9) * UpY
                    SAZ.Value = V(8) * UpZ
                End If
                Pose.F = V(10)
                Pose.L = V(11)
                If W(12) = "" Then FDX.Value = 1 Else FDX.Value = V(12)
                If W(13) = "" Then FDY.Value = 1 Else FDY.Value = V(13)
                If W(14) = "" Then FDZ.Value = 1 Else FDZ.Value = V(14)
                If W(15) = "" Then FAX.Value = 1 Else FAX.Value = V(15)
                If W(16) = "" Then FAY.Value = 1 Else FAY.Value = V(16)
                If W(17) = "" Then FAZ.Value = 1 Else FAZ.Value = V(17)
                If Pose.B < N Then
                    J = AnimUpdt(Pose.A).Animation.NodeAnimationCollection(Pose.B).KeyFrames.Count ' aantal keyframes
                    For I = 0 To J - 1
                        KF(Pose.B, I) = AnimUpdt(Pose.A).Animation.NodeAnimationCollection(Pose.B).KeyFrames(I) ' onthoudt keyframes
                    Next
                End If
                SE = True
                AnimPoseBone()
            End If
        End If
    End Sub

    Sub AnimFileLoad(F As String)

        ' laadt een beweging

        Dim N As Short

        If AI < 1 Then Exit Sub

        If File.Exists(F) Then
            N = AnimUpdt(AI).Animation.NodeAnimationCollection.Count - 1 ' for all bones
            For I = 0 To N
                AnimUpdt(AI).Animation.NodeAnimationCollection(I).KeyFrames.Clear() ' clear all keyframes
            Next
            FI += 1
            ParsFile(FI) = F
            ParsSkip = False
            TextParsRuns(ParsFile(FI))
            FI -= 1
        Else
            ' MsgBox("Animatiefile " & F & " niet gevonden")
        End If
    End Sub

    Sub AnimHelp_Rendering() Handles AnimHelp.Rendering

        ' deze procedure is een black box dat animaties laat werken

        If AI > 0 Then
            For I = 1 To AI
                If AE(I) <> 0 Then If AnimUpdt(I) IsNot Nothing Then AnimUpdt(I).Update(Stopwatch.GetTimestamp(), Stopwatch.Frequency)
                AnimBoneFollow(I)
                AnimBoneTarget()
                SceneCamera.Follow()
            Next
        End If
    End Sub

    Sub ModelTrans_Changed(sender As Object, e As EventArgs) Handles ModelTrans.Changed

        '

        Dim M As SharpDX.Matrix
        'Dim P As String
        Dim C1 As System.Drawing.Color
        Dim X As Short
        Dim Z As Short

        If Navigate And MediaScene.KeyPressed > 0 Then
            If AnimUpdt(AI) IsNot Nothing Then
                M = ModelTrans.ToMatrix
                ModelAnim(SI).Root.ModelMatrix = M ' verplaats model
                MP = M.TranslationVector.ToVector3D
                X = MP.X / 0.008
                Z = MP.Z / 0.008
                'C1 = MediaScene.FloorMap.GetPixel(127 + X, 127 + Z)
                MY = C1.G * 1.6 / 1000 - MP.Y
                'MY /= CS
                'MenuText3.Text = X & ", " & C1.G & ", " & Z & ", " & MY
                SceneCamera.Move() ' verplaats camera
            End If
        End If
    End Sub


    Sub ModelAnimKeyframe()

        ' importeer keyframe 

        ' W(0) = "keyframe"
        ' V(1) = bone index
        ' V(2) = key index
        ' V(3) = dx 
        ' V(4) = dy
        ' V(5) = dz
        ' V(6) = hx 
        ' V(7) = hy
        ' V(8) = hz
        ' V(9) = hw
        ' V(10) = time

        Dim B As Short
        Dim K As Keyframe
        Dim N As Short
        Dim T As Vector3

        If AI > 0 Then
            N = AnimUpdt(AI).Animation.NodeAnimationCollection.Count ' aantal botten
            For B = 0 To N - 1
                If LCase(AnimUpdt(AI).Animation.NodeAnimationCollection(B).Node.Name) = GetOldBone(W(1)) Then GoTo gevonden
            Next
            'MsgBox("bot " & W(1) & " niet gevonden")
            'ParsSkip = True
            Exit Sub
gevonden:
            N = AnimUpdt(AI).Animation.NodeAnimationCollection(B).KeyFrames.Count - 1 ' aantal keyframes
            If V(2) > N Then AnimUpdt(AI).Animation.NodeAnimationCollection(B).KeyFrames.Add(New Keyframe)
            T = MM(AI, B).TranslationVector
            K.BoneIndex = B
            K.Translation = New SharpDX.Vector3(V(3), V(4), V(5)) + T
            If W(9) = "" Then ' als er geen Q.W gevonden wordt dan is de rotatie in graden
                K.Rotation = EulerToQuaternion(New Vector3(V(6), V(7), V(8)))
            Else ' rotatie als quaternion
                K.Rotation = New SharpDX.Quaternion(V(6), V(7), V(8), V(9) * AC(AI))
            End If
            K.Scale = New SharpDX.Vector3(1, 1, 1)
            K.Time = V(10)
            AnimUpdt(AI).Animation.NodeAnimationCollection(B).KeyFrames(V(2)) = K
        Else
            MsgBox("bot kan niet bewegen")
            ParsSkip = True
        End If
    End Sub
    '
    Sub ModelAnimKeyframes()

        ' exporteert keyframes, eventeel met andere tijden

        ' W(0) = "keyframes"
        ' W(1) = filename
        ' V(2) = [start tijd]
        ' V(3) = [eind tijd]
        ' V(4) = [eerste]
        ' V(5) = [laatste]
        ' V(6) = [translatie]
        ' V(7) = [rotatie]

        Dim B As Short ' aantal botten
        Dim I As Short
        Dim J As Short
        Dim K As Keyframe
        Dim M As Short ' eerste keyframe
        Dim N As Short ' aantal keyframes
        Dim E As Vector3 ' rotatie in graden
        Dim R As SharpDX.Quaternion ' rotatie
        Dim S As String
        Dim T As Vector3 ' translatie

        If AI > 0 Then ' als laatst toegevoegde animatie groter dan 0 is
            If AnimUpdt(AI) IsNot Nothing Then ' als er een animatie gestart is
                W(1) = GetFileAddress(W(1)) ' bestand voor keyframes
                FileOpen(1, W(1), OpenMode.Output)
                If W(2) = "" Then V(2) = AnimUpdt(AI).Animation.StartTime
                If W(3) = "" Then V(3) = AnimUpdt(AI).Animation.EndTime
                If V(3) = 0 Then V(3) = 2
                If W(4) = "" Then M = 0 Else M = V(4) ' eerste keyframe
                PrintLine(1, "times, " & V(2) & ", " & V(3))
                B = AnimUpdt(AI).Animation.NodeAnimationCollection.Count ' aantal botten
                PrintLine(1, "' bones " & B)
                PrintLine(1, "' version " & Date.Now)
                For I = 0 To B - 1 ' voor alle botten
                    S = GetNewBone(AnimUpdt(AI).Animation.NodeAnimationCollection(I).Node.Name) ' botnaam
                    N = AnimUpdt(AI).Animation.NodeAnimationCollection(I).KeyFrames.Count
                    If W(5) <> "" And V(5) < N Then N = V(5) ' laatste keyframe
                    For J = M To N - 1 ' voor alle keyframes
                        K = AnimUpdt(AI).Animation.NodeAnimationCollection(I).KeyFrames(J) ' bestaand keyframe
                        If V(6) = 0 Then ' standaard
                            T = K.Translation * SC - MM(AI, I).TranslationVector * SC ' verschil translatie
                        ElseIf V(6) = 1 Then
                            T = (K.Translation) * SC ' geen verschil translatie
                        Else
                            T = (K.Translation) * 0 ' geen translatie
                        End If
                        R = K.Rotation
                        R.W *= AC(AI)
                        If V(7) > 0 Then E = QuaternionToEuler(R)
                        If N - 1 - M = 0 Then
                            K.Time = 0
                        Else
                            K.Time = (J - M) * V(3) / (N - 1 - M)
                        End If
                        If V(7) = 0 Then ' standaard; rotatie als quaternion
                            PrintLine(1, "keyframe, " & S & LB(J - M) & SF(T.X) & SF(T.Y) & SF(T.Z) &
                                  SF(R.X) & SF(R.Y) & SF(R.Z) & SF(R.W) & SF(K.Time))
                        Else ' rotatie in graden
                            PrintLine(1, "keyframe, " & S & LB(J - M) & SF(T.X) & SF(T.Y) & SF(T.Z) &
                                  SF(E.X) & SF(E.Y) & SF(E.Z) & "," & SF(K.Time))
                        End If
                    Next ' keyframe
                Next ' bot
                FileClose(1)
            End If
        End If
    End Sub

    Sub ModelAnimMerge()

        ' combineert een eerder geimporteerd object met een animatie

        ' W(0) = "merge"
        ' W(1) = bestandnaam

        W(1) = GetFileAddress(W(1))
        If File.Exists(W(1)) Then
            Using ObjectRead As New Importer ' laad een model
                ModelAnim(0) = ObjectRead.Load(W(1))
            End Using
            If ModelAnim(0).HasAnimation Then ' als het nieuwe model een animatie heeft
                ModelAnim(SI).Animations = ModelAnim(0).Animations ' combineer de animatie met het oude model
            Else ' als het model geen animatie heeft
                SceneModel.AddNode(ModelAnim(0).Root) ' combineer het nieuwe model met het oude model
                ModelAnim(SI).Root = SceneModel
            End If
        Else
            MsgBox("model " & W(1) & " niet gevonden")
        End If
    End Sub

    Sub ModelAnimMoveto()

        ' beweegt een bot naar een opgegeven punt

        ' W(0) = "moveto"
        ' V(1) = bot
        ' V(2) = x
        ' V(3) = y
        ' V(4) = z

        Dim D As Vector3
        Dim I As Short
        Dim J As Short
        Dim K As Keyframe
        Dim L As Short
        Dim N As Short
        Dim O As Vector3
        Dim Q As SharpDX.Quaternion
        'Dim T As Vector3
        Dim R As Vector3
        Dim S As String

        D = New Vector3(V(2), V(3), V(4))
        O = AnimUpdt(AI).Animation.NodeAnimationCollection(V(1)).Node.ModelMatrix.TranslationVector
        S = AnimUpdt(AI).Animation.NodeAnimationCollection(V(1)).Node.Parent.Name
        N = AnimUpdt(AI).Animation.NodeAnimationCollection.Count
        I = 0
        While I < N And AnimUpdt(AI).Animation.NodeAnimationCollection(I).Node.Name <> S
            I += 1
        End While
        N = AnimUpdt(AI).Animation.NodeAnimationCollection(I).KeyFrames.Count
        L = 0
        While L < 100 And AnimUpdt(AI).Animation.NodeAnimationCollection(V(1)).Node.TotalModelMatrix.TranslationVector.Z < D.Z
            L += 1
            AE(AI) = False
            For J = 0 To N - 1
                K = AnimUpdt(AI).Animation.NodeAnimationCollection(I).KeyFrames(J)
                R = QuaternionToEuler(K.Rotation)
                R.X += V(5)
                R.Y += V(6)
                R.Z += V(7)
                Q = EulerToQuaternion(R)
                K.Rotation = Q
                AnimUpdt(AI).Animation.NodeAnimationCollection(I).KeyFrames(J) = K
            Next
            AE(AI) = True
            'MenuText3.Text = R.X & ", " & AnimUpdt(AI).Animation.NodeAnimationCollection(V(1)).Node.TotalModelMatrix.TranslationVector.Z
            Wait(V(8))
            DoEvents()
        End While
        FileText.Text = "klaar"
    End Sub

    Sub ModelAnimPoses()

        ' exporteert houdingen

        ' W(0) = "poses"
        ' W(1) = filename
        ' W(2) = options

        Dim B As Short ' aantal botten
        Dim I As Short
        Dim K As Keyframe
        Dim Q As SharpDX.Quaternion
        Dim R As SharpDX.Vector3 ' rotatie
        Dim S As String
        Dim T As SharpDX.Vector3 ' translatie

        If AI > 0 Then
            If AnimUpdt(AI) IsNot Nothing Then
                W(1) = GetFileAddress(W(1))
                FileOpen(1, W(1), OpenMode.Output)
                B = AnimUpdt(AI).Animation.NodeAnimationCollection.Count ' aantal botten
                For I = 0 To B - 1 ' voor alle botten
                    S = AnimUpdt(AI).Animation.NodeAnimationCollection(I).Node.Name
                    S = GetNewBone(S)
                    If W(2) = "k" Then ' exporteer keyframe houding
                        K = AnimUpdt(AI).Animation.NodeAnimationCollection(I).KeyFrames(0)
                        T = K.Translation - MM(AI, I).TranslationVector
                        R = QuaternionToEuler(K.Rotation) - QuaternionToEuler(MatrixToQuaternion(MM(AI, I)))
                    ElseIf W(2) = "m" Then ' exporteer model matrix houding
                        T = MM(AI, I).TranslationVector
                        Q = MatrixToQuaternion(MM(AI, I))
                        If W(3) = "-1" Then
                            Q.W *= -1
                        End If
                        R = QuaternionToEuler(Q)
                    ElseIf W(2) = "t" Then ' exporteer total model matrix houding
                        T = MM(AI, I).TranslationVector
                        Q = MatrixToQuaternion(MM(AI, I))
                        If W(3) = "-1" Then
                            Q.W *= -1
                        End If
                        R = QuaternionToEuler(Q)
                    Else
                        MsgBox(W(2) & " is geen geldige optie")
                        FileClose(1)
                        ParsSkip = True
                        Exit Sub
                    End If
                    'PrintLine(1, "' " & S)
                    PrintLine(1, "pose" & LB(AI) & ", " & S & ", 2" &
                    SF(T.X) & SF(T.Y) & SF(T.Z) & SF(R.X) & SF(R.Y) & SF(R.Z))
                Next
                FileClose(1)
            End If
        End If
    End Sub

    Sub ModelAnimSync()

        ' synchroniseert een bot met een pose

        ' W(0) = "sync"
        ' V(1) = animatie index
        ' V(2) = bot nummer van model 2

        'Dim N As Short

        Sync.A = V(1)

        If W(2) = V(2).ToString Then ' als een bot nummer opgegeven is
            Sync.B = Bone(Pose.A, V(2)).I
        Else ' als een bot naam opgegeven is
            W(2) = GetOldBone(W(2))
            For I = 0 To 400
                If Bone(Sync.A, I).N = W(2) Then
                    Sync.B = Bone(Sync.A, I).I
                    GoTo verder
                End If
            Next
            MsgBox("bot " & W(2) & " niet gevonden")
            ParsSkip = True
            Exit Sub
        End If
verder:
    End Sub

    Sub ModelAnimTimes()
        '
        ' animatie tijden
        '
        ' W(0) = "times"
        ' V(1) = start tijd
        ' V(2) = eind tijd
        '
        If AI > 0 Then
            If AnimUpdt(AI) IsNot Nothing Then
                AnimUpdt(AI).Animation.StartTime = V(1)
                AnimUpdt(AI).Animation.EndTime = V(2)
            End If
        End If
    End Sub

    Sub ModelAnimWalk()

        ' navigeert een model door een omgeving

        If Navigate Then

            Dim O As Media3D.Vector3D ' model positie
            Dim P As New Media3D.Point3D(0, MY / SC, MZ / SC)

            Dim R As New Media3D.RotateTransform3D(New Media3D.AxisAngleRotation3D(New Media3D.Vector3D(0, 1, 0), MA))

            P = R.Transform(P)
            O = MP / 2 / SC ' ModelAnim(SI).Root.ModelMatrix.TranslationVector.ToVector3D / 2 '/ SC ' model positie

            Dim TTX As New Media3D.TranslateTransform3D(O)
            Dim DAX As New DoubleAnimation(0, P.X, New Windows.Duration(TimeSpan.FromSeconds(MT)))
            Dim TTY As New Media3D.TranslateTransform3D(O)
            Dim DAY As New DoubleAnimation(0, P.Y, New Windows.Duration(TimeSpan.FromSeconds(MT / 100)))
            Dim TTZ As New Media3D.TranslateTransform3D(O)
            Dim DAZ As New DoubleAnimation(0, P.Z, New Windows.Duration(TimeSpan.FromSeconds(MT)))
            '
            TTX.BeginAnimation(Media3D.TranslateTransform3D.OffsetXProperty, DAX)
            TTY.BeginAnimation(Media3D.TranslateTransform3D.OffsetYProperty, DAY)
            TTZ.BeginAnimation(Media3D.TranslateTransform3D.OffsetZProperty, DAZ)

            Dim S As New ScaleTransform3D(New Vector3D(SC, SC, SC))

            ModelTrans.Children(0) = R
            ModelTrans.Children(1) = TTX
            ModelTrans.Children(2) = TTY
            ModelTrans.Children(3) = TTZ
            ModelTrans.Children(4) = S
        End If
    End Sub

    Sub ModelAnimNavigate()

        ' start het navigeren van een model

        ' W(0) = "navigate"
        ' V(1) = model index
        ' V(2) = naar x als absolute waarde
        ' V(3) = naar Y
        ' V(4) = naar z
        ' V(5) = camera angle y
        ' V(6) = speed (m/s)

        Dim D As Media3D.Point3D ' delta afstand
        Dim O As Media3D.Point3D ' object positie

        Navigate = True
        MediaScene.Focus()

        If W(1) <> "" Then
            If V(1) < 0 Then ' stop navigeren
                Navigate = False
            Else
                O = ModelAnim(V(1)).Root.ModelMatrix.TranslationVector.ToPoint3D
                D.X = V(2) - O.X ' model is in O en moet D overbruggen
                D.Y = V(3) - O.Y ' hoogte gebeurd niets mee
                D.Z = V(4) - O.Z
                If V(5) <> 0 Then SceneCamera.CA = V(5) ' camera hoek
                MA = Math.Atan(D.X / D.Z) * (180 / Math.PI) ' bereken model hoek
                If D.Z < 0 Then MA += 180
                MZ = (D.X ^ 2 + D.Z ^ 2) ^ 0.5 ' afstand
                'MY = D.Y
                If V(6) = 0 Then V(6) = 1.38 ' snelheid in m/s
                MT = MZ / V(6) ' tijd
                ModelAnimWalk()
            End If
        Else
            MY = 0
            MZ = 0
            AnimFileLoad(IdleFile)
        End If
        SceneModelShow()
    End Sub

    Sub SAX_ValueChanged() Handles SAX.ValueChanged

        ' als slider verschoven wordt

        AnimPoseBone()
    End Sub

    Sub SAY_ValueChanged() Handles SAY.ValueChanged

        ' als slider verschoven wordt

        AnimPoseBone()
    End Sub

    Sub SAZ_ValueChanged() Handles SAZ.ValueChanged

        ' als slider verschoven wordt

        AnimPoseBone()
    End Sub

    Sub SDX_ValueChanged() Handles SDX.ValueChanged

        ' als slider verschoven wordt

        AnimPoseBone()
    End Sub

    Sub SDY_ValueChanged() Handles SDY.ValueChanged

        ' als slider verschoven wordt

        AnimPoseBone()
    End Sub

    Sub SDZ_ValueChanged() Handles SDZ.ValueChanged

        ' als slider verschoven wordt

        AnimPoseBone()
    End Sub

    Sub FDX_ValueChanged() Handles FDX.ValueChanged

        ' als slider verschoven wordt

        AnimPoseBone()
    End Sub

    Sub FDY_ValueChanged() Handles FDY.ValueChanged

        ' als slider verschoven wordt

        AnimPoseBone()
    End Sub

    Sub FDZ_ValueChanged() Handles FDZ.ValueChanged

        ' als slider verschoven wordt

        AnimPoseBone()
    End Sub

    Sub FAX_ValueChanged() Handles FAX.ValueChanged

        ' als slider verschoven wordt

        AnimPoseBone()
    End Sub

    Sub FAY_ValueChanged() Handles FAY.ValueChanged

        ' als slider verschoven wordt

        AnimPoseBone()
    End Sub

    Sub FAZ_ValueChanged() Handles FAZ.ValueChanged

        ' als slider verschoven wordt

        AnimPoseBone()
    End Sub

    Sub TB_Textchanged() Handles TB.TextChanged
        '
        '
        '
        Dim J As Short
        Dim N As Short
        '
        If ModelAnim(SI) IsNot Nothing Then
            If ModelAnim(SI).HasAnimation Then
                If Right(TB.Text, 1) = "+" Then
                    J = Val(TB.Text)
                    N = AnimUpdt(AI).Animation.NodeAnimationCollection.Count - 1
                    If J > N Then
                        TB.Text = N
                    End If
                    SDX.Value = 0
                    SDY.Value = 0
                    SDZ.Value = 0
                    SAX.Value = 0
                    SAY.Value = 0
                    SAZ.Value = 0
                End If
            End If
        End If
    End Sub

    Sub AnimBoneFollow(A As Short)

        ' laat een onderdeel een bot van een model volgen

        Dim D As Vector3
        Dim I As Short
        Dim M As SharpDX.Matrix
        '
        For I = 1 To 10
            If FB(A, I).I = 0 Or FB(A, I).B > 399 Then Exit Sub
            M = FB(A, I).T.ToMatrix * AnimUpdt(A).Animation.NodeAnimationCollection(FB(A, I).B).Node.TotalModelMatrix
            SceneModel.SceneNode.Items(FB(A, I).I).ModelMatrix = M
            SceneModel.SceneNode.Items(FB(A, I).I).ComputeTransformMatrix()
            D = M.TranslationVector
            TQ.Text = "volg " & SF(D.X) & SF(D.Y) & SF(D.Z)
        Next
    End Sub

    Sub ModelAnimBone()

        ' laadt bot

        ' W(0) = "bone"
        ' V(1) = index
        ' V(2) = nieuwe index
        ' W(3) = oude naam
        ' W(4) = nieuwe naam

        Bone(AI, V(1)).I = V(2)
        Bone(AI, V(1)).O = LCase(W(3))
        Bone(AI, V(1)).N = LCase(W(4))
        BoneIndex(AI) += 1
    End Sub

    Sub ModelAnimBones()

        ' exporteert de botten van een model naar de bottenlijst en eventueel naar een bestand

        ' W(0) = "bones"
        ' W(1) = [bestandnaam]

        Dim N As Short
        Dim S As String

        If ModelAnim(SI).HasAnimation Then
            If W(1) <> "" Then ' als een bestandnaam is opgegeven
                W(1) = GetFileAddress(W(1))
                FileOpen(1, W(1), OpenMode.Output)
                PrintLine(1, "' bones " & N)
                PrintLine(1, "' version " & Now.ToString)
            End If
            BoneList.Items.Clear()
            N = ModelAnim(SI).Animations(0).NodeAnimationCollection.Count ' aantal botten
            For I = 0 To N - 1
                S = LCase(ModelAnim(SI).Animations(0).NodeAnimationCollection(I).Node.Name)
                BoneList.Items.Add(I & ", " & Bone(SI, I).N)
                If W(1) <> "" Then PrintLine(1, "bone, " & I & ", " & I & ", " & Bone(SI, I).O & ", " & Bone(SI, I).N)
            Next
            FileClose(1)
        Else
            MsgBox("model heeft geen botten")
        End If
    End Sub

    Sub AnimBoneFollow()

        ' laat een onderdeel een bot volgen

        ' W(0) = "follow"
        ' V(1) = volg index
        ' V(2) = [onderdeel index]
        ' W(3) = bot als naam of als index
        ' V(4) = dx offset x
        ' V(5) = dy
        ' V(6) = dz
        ' V(7) = hx hoek x
        ' V(8) = hy
        ' V(9) = hz
        ' V(10) = sx schaal x
        ' V(11) = sy
        ' V(12) = sz

        Dim I As Short

        If W(2) <> "" Then
            FB(AI, V(1)).I = V(2)
        Else
            FB(AI, V(1)).I = SceneModel.SceneNode.ItemsCount - 1
        End If

        If (W(3)) <> V(3).ToString Then ' als de naam van een bot is opgegeven
            I = 0
            While Bone(AI, I).N <> W(2) ' zoek het botnummer
                I += 1
            End While
            V(3) = I
        End If
        FB(AI, V(1)).B = V(3) ' onthoud bot nummer
        FB(AI, V(1)).T = New Transform3DGroup
        ModelPart.Translate(FB(AI, V(1)).T, V(4), V(5), V(6))
        ModelPart.Rotate(FB(AI, V(1)).T, New Point3D(0, 0, 0), V(7), V(8), V(9))
        If W(10) = "" Then V(10) = 1
        If W(11) = "" Then V(11) = 1
        If W(12) = "" Then V(12) = 1
        ModelPart.Scale(FB(AI, V(1)).T, New Vector3D(V(10), V(11), V(12)))
    End Sub

    Sub AnimPoseBone()

        Dim A As Double
        Dim I As Short
        Dim K As Keyframe
        Dim K2 As Keyframe
        Dim N As Short
        Dim Q As SharpDX.Quaternion
        Dim R As Vector3
        Dim D As Vector3
        Dim F As Vector3
        Dim S As String
        Dim T As Vector3

        If SE Then
            N = AnimUpdt(Pose.A).Animation.NodeAnimationCollection.Count
            If Pose.B < N Then
                T = New Vector3(SDX.Value, SDY.Value, SDZ.Value)
                R = New Vector3(SAX.Value, SAY.Value, SAZ.Value)
                D = New Vector3(FDX.Value, FDY.Value, FDZ.Value)
                F = New Vector3(FAX.Value, FAY.Value, FAZ.Value)
                N = AnimUpdt(Pose.A).Animation.NodeAnimationCollection(Pose.B).KeyFrames.Count ' aantal keyframes
                If N < 5 Then

                End If
                If Pose.L > 0 Then N = Pose.L ' laatste keyframe
                If Pose.F = 0 Then Pose.F = 1 ' eerste keyframe
                S = AnimUpdt(Pose.A).Animation.NodeAnimationCollection(Pose.B).Node.Name
                For I = Pose.F - 1 To N - 1 ' voor alle keyframes
                    K = KF(Pose.B, I)
                    If Pose.T = 0 Then ' geen beweging
                        K.Translation = MM(Pose.A, Pose.B).TranslationVector
                        K.Rotation = New SharpDX.Quaternion(0, 0, 0, AC(Pose.A))
                    ElseIf Pose.T = 1 Then ' alleen translatie
                        K.Translation = MM(Pose.A, Pose.B).TranslationVector + T
                        K.Rotation = New SharpDX.Quaternion(0, 0, 0, AC(Pose.A))
                    ElseIf Pose.T = 2 Then ' corrigeer keyframes
                        K.Translation = K.Translation * D + T

                        Q = EulerToQuaternion(R)
                        Q.W *= AC(Pose.A)
                        K.Rotation *= Q
                    ElseIf Pose.T = 3 Then ' heen en weer
                        A = 180 * (I - (Pose.F - 1)) / ((N - 1) - (Pose.F - 1))
                        Q = EulerToQuaternion(R * DYS(1, A))
                        Q.W *= AC(Pose.A)
                        K.Rotation *= Q
                        K.Translation += T
                    ElseIf Pose.T = 4 Then ' slingeren
                        A = 360 * I / (N - 1)
                        Q = EulerToQuaternion(R * DYS(1, A))
                        Q.W *= AC(Pose.A)
                        K.Rotation *= Q
                        K.Translation += T
                    End If
                    AnimUpdt(Pose.A).Animation.NodeAnimationCollection(Pose.B).KeyFrames(I) = K
                    If Sync.B > -1 Then
                        K2 = K
                        K2.Translation = AnimUpdt(Sync.A).Animation.NodeAnimationCollection(Sync.B).Node.ModelMatrix.TranslationVector
                        AnimUpdt(Sync.A).Animation.NodeAnimationCollection(Sync.B).KeyFrames(I) = K2
                    End If
                Next
                TA.Text = "pose, " & Pose.A & ", " & Pose.B & ", " & Pose.T & SF(SDX.Value) & SF(SDY.Value) & SF(SDZ.Value) &
                                                                          SF(SAX.Value) & SF(SAY.Value) & SF(SAZ.Value)
                TB.Text = "corr, " & Pose.A & ", " & Pose.B & ", " & Pose.T & SF(FDX.Value) & SF(FDY.Value) & SF(FDZ.Value) &
                                                                          SF(FAX.Value) & SF(FAY.Value) & SF(FAZ.Value)
            End If
        End If
    End Sub

    Sub AnimBoneTarget()

        Dim D As Vector3

        D = AnimUpdt(AI).Animation.NodeAnimationCollection(BoneTarget).Node.ModelMatrix.TranslationVector
        TT.Text = "target " & SF(D.X) & SF(D.Y) & SF(D.Z)
    End Sub

    Function GetNewBone(O As String) As String

        ' zoekt nieuwe bot naam

        Dim I = 0

        O = LCase(O)

        While Bone(AI, I).O <> O And I < BoneIndex(AI)
            I += 1
        End While

        If I < BoneIndex(AI) Then
            GetNewBone = Bone(AI, I).N
        Else
            GetNewBone = O
        End If
    End Function

    Function GetOldBone(N As String) As String

        ' zoekt oude bot naam

        Dim I = 0

        N = LCase(N)

        While Bone(AI, I).N <> N And I < BoneIndex(AI)
            I += 1
        End While

        If I < BoneIndex(AI) Then
            GetOldBone = Bone(AI, I).O
        Else
            GetOldBone = N ' houdt nieuwe naam
        End If
    End Function

    Sub BoneList_SelectionChanged() Handles BoneList.SelectionChanged

        ' selecteert een bot

        Dim I As Short
        Dim J As Short
        Dim S As String

        If BoneList.SelectedItem = Nothing Then Exit Sub
        S = BoneList.SelectedItem
        I = InStr(S, ",") - 1
        I = Val(Left(S, I))
        Pose.B = I
        J = AnimUpdt(Pose.A).Animation.NodeAnimationCollection(Pose.B).KeyFrames.Count ' aantal keyframes
        For I = 0 To J - 1
            KF(Pose.B, I) = AnimUpdt(Pose.A).Animation.NodeAnimationCollection(Pose.B).KeyFrames(I) ' onthoudt keyframes
        Next
        SAX.Value = 0
        SAY.Value = 0
        SAZ.Value = 0
        AnimPoseBone()
    End Sub

    Public Structure TBone
        Dim I As Short  ' index
        Dim O As String ' oude naam
        Dim N As String ' nieuwe naam
    End Structure

    Public Structure TFollow ' type voor volgen van een bot door een onderdeel
        Dim B As Short ' bot nummer
        Dim I As Short ' onderdeel nummer
        Dim T As Transform3DGroup
    End Structure

    Public Structure TPose ' type houding
        Dim A As Short ' animatie index
        Dim B As Short ' bot index
        Dim T As Short ' animatie type
        Dim F As Short ' begin keyframe
        Dim L As Short ' eind keyframe
    End Structure

    Public Structure TSync
        Dim A As Short
        Dim B As Short
    End Structure


    Public WithEvents ModelTrans As New Transform3DGroup
    Public WithEvents AnimHelp As New CompositionTargetEx

    Public FB(4, 10) As TFollow ' volg bot
    Public SI As Short ' model index
    Public AI As Short ' animatie index
    Public AE(4) As Boolean ' animatie enabled
    Public AC(4) As Short ' quaternion.w factor
    Public ATX(4) As Short ' animatie totaal voor x
    Public ATY(4) As Short ' animatie totaal voor y
    Public ATZ(4) As Short ' animatie totaal voor z
    Public IdleFile As String
    Public WalkFile As String
    Public Sync As TSync
    Public Navigate As Boolean ' niet navigeren
    Public MA As Short ' model hoek met y as
    Public MZ As Single ' model delta verplaatsing langs z-as
    Public MY As Single ' model delta verplaatsing langs y-As
    Public MT As Single ' model tijd voor verplaatsing
    Public MP As Vector3D ' model positie
    Public Bone(4, 400) As TBone
    Public BoneIndex(4) As Short
    Public Paint As Boolean
    Public Up As Char = "Y"
    Public UpX As Short = 1
    Public UpY As Short = 1
    Public UpZ As Short = 1
    Public BoneTarget As Short
    Public BoneMatrix As SharpDX.Matrix
    Public KF(400, 1000) As Keyframe
    Public Pose As TPose
    Public WithEvents BoneList As New ListBox
    Public WithEvents SDX As New Slider
    Public WithEvents SDY As New Slider
    Public WithEvents SDZ As New Slider
    Public WithEvents SAX As New Slider
    Public WithEvents SAY As New Slider
    Public WithEvents SAZ As New Slider
    Public WithEvents FDX As New Slider
    Public WithEvents FDY As New Slider
    Public WithEvents FDZ As New Slider
    Public WithEvents FAX As New Slider
    Public WithEvents FAY As New Slider
    Public WithEvents FAZ As New Slider
    Public WithEvents TB As New TextBox
    Public SE As Boolean ' slide enabled
    Public TA As New TextBox
    Public TH As New TextBox
    Public TK As New TextBox
    Public TN As New TextBox
    Public TT As New TextBox
    Public TQ As New TextBox
    Public TS As New TextBox ' sync
    Public TX As New TextBox '
    Public AnimUpdt(4) As NodeAnimationUpdater
    Public MM(4, 400) As SharpDX.Matrix ' model matrix na laden van model
    Public TM(4, 400) As SharpDX.Matrix ' model matrix na laden van model
End Module
