Imports SharpDX
Imports System.IO
Imports System.Windows.Input
Imports System.Windows.Threading

Module UnitTextPars

    ' van UnitText

    Sub TextParsInit()

        ' start het uitvoeren van opdrachten

        ParsSkip = False
        PhotoName = ""
        ParsWord = ""
        GridText.Save()
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
        FI = 0
        PN = 0
        UN = 0
        NN = 0
        Log = ""
        MI = 0
        For I = 0 To 20
            INV(I) = False
            REV(I) = False
        Next
        MouseCursor = Mouse.OverrideCursor ' onthoudt textcursor
        Mouse.OverrideCursor = Cursors.Wait
        Mouse.UpdateCursor()

        TextParsRuns(TextFile)
        TextParsDone()
    End Sub

    Sub TextParsDone()

        ' stopt het uitvoeren van opdrachten

        PartLegoDone()
        If PartData Then PartDataList()
        DoEvents()
        If PhotoName <> "" Then
            Wait(1)
            GridMedia.Photo()
        End If
        Mouse.OverrideCursor = Cursors.Arrow
        Mouse.UpdateCursor()
    End Sub

    Sub TextParsRuns(S As String)

        ' voert de opdrachten uit

        Dim A() As String ' tekst array
        Dim C As Char
        Dim K As Short
        Dim N As Long
        Dim P As Long
        Dim T As String

        If File.Exists(S) Then
            ParsFile(FI) = S
            A = File.ReadAllLines(ParsFile(FI))
        Else
            MsgBox("bestand " & vbCrLf & ParsFile(FI) & vbCrLf & "in" & vbCrLf & ParsFile(FI - 1) & vbCrLf & "niet gevonden")
            ParsSkip = True
            Exit Sub
        End If
        T = GetFileType(ParsFile(FI))
        If T = ".bvh" Or T = ".dae" Or T = ".obj" Then
            SceneModelInit()
            W(1) = ParsFile(FI)
            W(2) = "0"
            V(2) = 0
            SceneModelImport()
            Exit Sub
        End If
        If T = ".dat" Or T = ".mtl" Then
            C = " "
        Else
            C = ","
        End If
        N = UBound(A)
        If ParsTo Then
            K = GridText.GetLineIndexFromCharacterIndex(GridText.CaretIndex)
            If N > K Then N = K
        End If
        For P = 0 To N ' pars alle regels
            While WaitKey
                DoEvents()
            End While
            If JumpWord = "" Then
                TextParsLine(ParsWord & A(P), C)
                If ParsWord <> "" And W(1) = "pars" Then
                    ParsWord = ""
                Else
                    TextParsWord()
                End If
                If ParsSkip Then
                    Exit Sub
                End If
            Else
                If A(P) = JumpWord Then JumpWord = ""
            End If
            'DoEvents()
        Next
        ParsWord = ""
    End Sub

    Sub TextParsFile()

        ' voegt een bestand toe aan de uitvoering van de opdrachten

        ' W(0) = "file"
        ' W(1) = filename
        ' V(2) = [dx]
        ' V(3) = [dy]
        ' V(4) = [dz]
        ' V(5) = [ax]
        ' V(6) = [ay]
        ' V(7) = [az]

        Dim T As String

        T = GetFileAddress(W(1))

        FI += 1

        ParsFile(FI) = T

        OX(FI) = V(2) * SC * GS
        OY(FI) = V(3) * SC * GS
        OZ(FI) = V(4) * SC * GS

        AX(FI) = V(5)
        AY(FI) = V(6)
        AZ(FI) = V(7)

        If W(8) = "" Then SX(FI) = 1 Else SX(FI) = V(8)
        If W(9) = "" Then SY(FI) = 1 Else SY(FI) = V(9)
        If W(10) = "" Then SZ(FI) = 1 Else SZ(FI) = V(10)

        T = GetFileType(T)
        If T = ".asm" Then

        ElseIf T = ".dat" Then
            PartFile = GetFileName(W(1))
            TextParsRuns(ParsFile(FI))
        ElseIf T = ".bin" Then
            PartLegoBin(ParsFile(FI))
            PartLegoAdd()
        Else ' .txt
            If ParsSkip Then Exit Sub
            If W(2) <> "" And W(3) = "" Then OY(FI) = StepHeight * SC * GS
            TextParsRuns(ParsFile(FI))
        End If
        FI -= 1
    End Sub

    Sub TextParsStep()

        ' voert 1 opdracht uit

        Dim I As Short
        Dim N As Short

        N = Len(GridText.Text) - 1
        I = GridText.GetLineIndexFromCharacterIndex(GridText.CaretIndex)
        While GridText.GetLineIndexFromCharacterIndex(GridText.CaretIndex) = I And GridText.CaretIndex < N
            GridText.CaretIndex += 1
        End While

        TextParsInit()
    End Sub

    Sub TextParsTags()

        ' leest de trefwoorden uit een afbeelding bestand

        Dim B() As Byte
        Dim I As Short
        Dim S As String = ""

        W(1) = GetFileAddress(W(1))
        B = File.ReadAllBytes(W(1))
        I = &HFF
        While B(I) < &HDF
            If B(I) > 19 Then
                If Chr(B(I)) = ";" Then
                    S &= vbCrLf
                Else
                    S &= Chr(B(I))
                End If
            End If
            I += 1
        End While
        MsgBox(S)
    End Sub

    Sub TextParsJump()

        JumpWord = W(1) & ":"
    End Sub

    Sub TextParsLine(S As String, C As Char)

        ' leest een regel

        Dim I As Short
        Dim J As Short
        Dim K As Short

        For I = 0 To 20
            W(I) = ""
            V(I) = 0
        Next

        J = 0
        S = Trim(S)
        Do
            If S <> "" Then
                If S(0) = """" Then
                    S = Mid(S, 2)
                    K = InStr(S, """")
                    S = Left(S, K - 1) & Mid(S, K + 1)
                Else
                    K = InStr(S, C) ' zoek volgende scheidings karakter
                End If
                If K > 0 Then ' als er een scheidings karakter is
                    W(J) = Trim(Left(S, K - 1)) ' isoleer woord
                    S = Trim(Mid(S, K + 1))
                Else ' als er scheidings karakter meer is 
                    W(J) = Trim(S) ' isoleer laatste woord
                End If
                J += 1
            Else
                K = 0
            End If
        Loop Until K = 0 Or J = 20

        Try
            For I = 0 To 20
                V(I) = Val(W(I))
            Next
        Catch ex As Exception

        End Try
    End Sub

    Sub TextParsNext()

        ' voegt een woord toe aan de daarop volgende opdrachten

        ' W(0) = "pars"
        ' W(1) = word

        Dim I As Short

        ParsWord = W(1) & ", "
        I = 2
        While W(I) <> ""
            ParsWord &= W(I)
            I += 1
        End While
    End Sub

    Sub TextParsOpen()

        ' opent een bestand en voert de opdrachten uit

        ' W(0) = "open"
        ' W(1) = bestand

        TabsFile.Load(GetFileAddress(W(1)))
        If Not ParsSkip Then TextParsInit()
    End Sub

    Sub TextParsWord()

        ' zoekt de opdracht bij een woord

        Select Case LCase(W(0))

            Case "0"
                PartLegoComment()
            Case "1"
                PartLegoFile()
            Case "3"
                PartLegoTriangle()
            Case "4"
                PartLegoRectangle()

            Case "albedocolor"
                PBR.AlbedoColor = New Color4(V(1), V(2), V(3), V(4))
            Case "albedomap"
                PBR.AlbedoMapFilePath = W(1)
                PBR.AlbedoMap = LoadMapsToMemory(PBR.AlbedoMapFilePath)
            Case "album"
                PlayMakeAlbum()
            Case "ambientmap"
                PBR.AmbientOcculsionMapFilePath = W(1)
                PBR.AmbientOcculsionMap = LoadMapsToMemory(W(1))
            Case "angle"
                MediaDraw.Angle()
            Case "animate"
                ModelAnimStart()
            Case "artist"
                PlayMakeArtist()

            Case "back"
                MediaDraw.Back()
            Case "backup"
                FileMakeBackup()
            Case "bar"
                PlayMidiBar()
            Case "block"
                TabsCtrl.ParsBlock()
            Case "bone"
                ModelAnimBone()
            Case "bones"
                ModelAnimBones()
            Case "box"
                ModelPart.Box()
            Case "button"
                TabsCtrl.ParsButton()

            Case "camera"
                SceneCamera.Pars()
            Case "circle"
                MediaDraw.Circle()
            Case "clear"
                MediaDrawClear()
            Case "column"
                DataTableColumn()
            Case "control"
                TabsCtrl.Runs()
            Case "cylinder"
                ModelPart.Cylinder()

            Case "d", "tr"
                MaterialTrans = V(1)
            Case "data"
                DataTableData()
            Case "database"
                MediaData.Init()
            Case "debug"
                MsgBox("debug")
            Case "dirs"
                DataTableDirs()
            Case "draw"
                MediaDraw.Runs()

            Case "export"
                SceneModelExport()

            Case "face", "f"
                ModelPartFace()
            Case "field"
                ParsGameField()
            Case "file"
                TextParsFile()
            Case "files"
                DataTableFiles()
            Case "floor"
                MediaScene.Floor()
            Case "follow"
                AnimBoneFollow()

            Case "game"
                ParsGameFile()
            Case "grid"
                MediaDrawGrid()

            Case "idle"
                IdleFile = GetFileAddress(W(1))
            Case "if"
                DataTableIf()
            Case "import"
                SceneModelImport()
            Case "info"
                DataTableInfo()
            Case "input"
                TabsCtrl.ParsInput()

            Case "jump"
                TextParsJump()

            Case "ka"
                PartMaterial.AmbientColor = New SharpDX.Color4(V(1), V(2), V(3), MaterialTrans)
            Case "kd"
                PartMaterial.DiffuseColor = New SharpDX.Color4(V(1), V(2), V(3), MaterialTrans)
            Case "ke"
                PartMaterial.EmissiveColor = New SharpDX.Color4(V(1), V(2), V(3), MaterialTrans)
            Case "kr"
                PartMaterial.ReflectiveColor = New SharpDX.Color4(V(1), V(2), V(3), MaterialTrans)
            Case "ks"
                PartMaterial.SpecularColor = New SharpDX.Color4(V(1), V(2), V(3), MaterialTrans)
            Case "keep"
                ModelPartKeep()
            Case "key"
                WaitOnKey()
            Case "keyframe"
                ModelAnimKeyframe()
            Case "keyframes"
                ModelAnimKeyframes()

            Case "lamp"
                SceneLamp.Pars()
            Case "letters"
                ModelPartLetters()
            Case "line"
                MediaDrawLine()
            Case "list"
                PartData = True
            Case "load"
                TabsCtrl.Load()
            Case "log"
                Log = W(1)

            Case "makedirs"
                FileMakeDirs()
            Case "makefile"
                FileMakeFile()
            Case "map"
                ModelPartMap()
            Case "maps"
                ModelPartMaps()
            Case "map_bump"
                PartMaterial.DisplacementMapFilePath = W(1)
                PartMaterial.DisplacementMap = LoadMapsToMemory(W(1))
                'PartMaterial.NormalMapFilePath = W(1)
                'PartMaterial.NormalMap = LoadMapsToMemory(W(1))
                PartMaterial.DisplacementMapScaleMask = New Vector4(0.2F, 0.2F, 0.2F, 0)
                PartMaterial.EnableTessellation = True
                PartMaterial.MaxDistanceTessellationFactor = 1
                PartMaterial.MinDistanceTessellationFactor = 3
                PartMaterial.MaxTessellationDistance = 500
                PartMaterial.MinDistanceTessellationFactor = 2
                PartMaterial.EnableAutoTangent = True
            Case "map_d"
                PartMaterial.DiffuseAlphaMapFilePath = W(1)
                PartMaterial.DiffuseAlphaMap = LoadMapsToMemory(W(1))
            Case "map_kd"
                PartMaterial.DiffuseMapFilePath = W(1)
                PartMaterial.DiffuseMap = LoadMapsToMemory(W(1))
                'PartMaterial.DiffuseAlphaMapFilePath = W(1)
                'PartMaterial.DiffuseAlphaMap = LoadMapsToMemory(W(1))
            Case "map_kn"
                PartMaterial.NormalMapFilePath = W(1)
                PartMaterial.NormalMap = LoadMapsToMemory(W(1))
            Case "map_ks"
                PartMaterial.SpecularColorMapFilePath = W(1)
                PartMaterial.SpecularColorMap = LoadMapsToMemory(W(1))
            Case "material"
                PartMaterial.Import()
            Case "materials"
                PartMaterial.Export()
            Case "matrix"
                SceneModelMatrix()
            Case "matrices"
                SceneModelMatrices()
            Case "merge"
                ModelAnimMerge()
            Case "midi"
                PlayMidiStart()
            Case "mode"
                FormGrid.Pars()
            Case "move"
                ModelPartMove()
            Case "moveto"
                ModelAnimMoveto()

            Case "navigate"
                ModelAnimNavigate()
            Case "new", "g"
                ModelPartNew()
            Case "newmtl"
                'PartMaterial.Name = W(1)
            Case "normal", "vn"
                ModelPartNormal()
            Case "nodes"
                SceneModelNodes()
            Case "note"
                PlayMidiNote()
            Case "ns"
                PartMaterial.SpecularShininess = V(1)

            Case "offset"
                MediaDrawOffset()
            Case "open"
                TextParsOpen()
            Case "output"
                TabsCtrl.ParsOutput()

            Case "pars"
                TextParsNext()
            Case "part"
                ModelPartImport()
            Case "parts"
                ModelPartExport()
            Case "photo"
                If W(1) = "" Then W(1) = "\info.jpg"
                PhotoName = GetFileAddress(W(1))
            Case "pict"
                MediaDrawPicture()
            Case "pillar"
                ModelPartPillar()
            Case "play"
                MediaDrawPlay()
            Case "point", "v"
                ModelPartPoint()
            Case "points"
                ModelPartPoints()
            Case "pose"
                ModelAnimPose()
            Case "poses"
                ModelAnimPoses()
            Case "pyramid"
                ModelPart.Pyramid()

            Case "rectangle"
                MediaDrawRectangle()
            Case "reform"
                ModelPartReform()
            Case "row"
                DataTableRows()
            Case "rung"
                TabsCtrl.ParsRung()

            Case "say"
                TextTalk.Say()
            Case "scale"
                MediaDrawScale()
            Case "scene"
                MediaScene.Runs()
            Case "segment"
                ModelPart.Segment()
            Case "shadow"
                MediaScene.Shadow()
            Case "skip"
                TextParsSkip()
            Case "slope"
                ModelPart.Slope()
            Case "sort"
                DataTableSort()
            Case "sphere"
                ModelPart.Sphere()
            Case "spin"
                MediaScene.Spin()
            Case "start"
                TextParsStart()
            Case "step"
                PartLegoStep()
            Case "sync"
                ModelAnimSync()

            Case "table"
                DataTableAdd()
            Case "tags"
                TextParsTags()
            Case "talk"
                TextTalk.Init()
            Case "target"
                BoneTarget = V(1)
                BoneMatrix.TranslationVector = New Vector3(V(2), V(3), V(4))
            Case "tempo"
                PlayMidiTempo()
            Case "test"
                TabsTest.Test()
            Case "text"
                MediaDrawText()
            Case "texture", "vt"
                ModelPartTexture()
            Case "timer"
                TabsCtrl.ParsTimer()
            Case "times"
                ModelAnimTimes()
            Case "track"
                PlayFindTrack() ' zoek muziek bestand
            Case "train"
                TabsCtrl.ParsTrain()
            Case "triangle"
                MediaDrawTriangle()
            Case "turn"
                ModelPartTurn()

            Case "up"
                Up = W(1)
                UpX = V(2)
                UpY = V(3)
                UpZ = V(4)
            Case "uvmap"
                GridMapsShow()

            Case "voice"
                PlayMidiVoice()

            Case "wait"
                Wait(V(1))
            Case "walk"
                WalkFile = GetFileAddress(W(1))
            Case "wall"
                ModelPartWall()
            Case "week"
                MsgBox(GetWeek(W(1)))
            Case "wires"
                SceneModelWires()

            Case "", "#", "2", "5", "l", "mtllib", "o", "s", "usemtl"
                ' overslaan
            Case Else
                If Left(W(0), 1) <> "'" And Left(W(0), 1) <> "-" Then
                    MsgBox("opdracht " & W(0) & " in " & vbCrLf & ParsFile(FI) & vbCrLf & " is niet herkent")
                    ParsSkip = True
                End If
        End Select
    End Sub

    Sub ParsGameField()

        GameField(V(1)).P.X = V(2)
        GameField(V(1)).P.Y = V(3)
        GameField(V(1)).P.Z = V(4)
        GameField(V(1)).T = V(5)
    End Sub

    Sub ParsGameFile()

        ' harten
        ' ruiten
        ' schoppen
        ' klaveren

        Dim I As Short
        Dim J(52) As Short
        Dim K As Short

        Dim Card() As String = {"", "H2", "H3", "H4", "H5", "H6", "H7", "H8", "H9", "H10", "HB", "HV", "HH", "HA",
                                "R2", "R3", "R4", "R5", "R6", "R7", "R8", "R9", "R10", "RB", "RV", "RH", "RA",
                                "S2", "S3", "S4", "S5", "S6", "S7", "S8", "S9", "S10", "SB", "SV", "SH", "SA",
                                "K2", "K3", "K4", "K5", "K6", "K7", "K8", "K9", "K10", "KB", "KV", "KH", "KA"}
        Game = True
        Randomize()
        For I = 1 To 52
            J(I) = (53 - I) * Rnd() + 1
            W(1) = Card(J(I)) & ".jpg"
            PartMaterial.Import()

            V(1) = 0.03
            V(2) = 0.04
            V(3) = 0.001
            V(4) = GameField(I).P.X * 0.04 - 0.16
            V(5) = 0.06 - GameField(I).P.Y * 0.01
            V(6) = GameField(I).P.Z * 0.001
            W(13) = Card(J(I))
            GameField(I).N = W(13)
            ModelPart.Box()
            For K = J(I) To 51
                Card(K) = Card(K + 1)
            Next
        Next
    End Sub

    Sub TextParsSkip()

        ' stopt opdrachten

        ParsSkip = True
    End Sub

    Sub TextParsStart()

        ' start een programma

        ' W(0) = "start"
        ' W(1) = programmapad

        Try
            Shell(W(1), AppWinStyle.NormalFocus)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Sub Wait(D As Single)

        ' wacht d secondes

        Dim S As Long ' start ticks

        S = Now.Ticks
        Do
            FileText.Text = "wacht " & Int(D - (Now.Ticks - S) / 10000000)
            DoEvents()
        Loop Until (Now.Ticks - S) / 10000000 > D
        FileText.Text = ""
    End Sub

    Sub WaitOnKey()

        ' wacht op het indrukken van een toets

        GridText.Focus()
        WaitKey = True
    End Sub

    Sub DoEvents()

        ' deblokkert andere opdrachten

        Dim Frame = New DispatcherFrame()

        Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, New DispatcherOperationCallback(AddressOf ExitFrame), Frame)
        Windows.Threading.Dispatcher.PushFrame(Frame)
    End Sub

    Function ExitFrame(F As Object) As Object

        ' hoort bij doevents

        F.Continue = False

        Return vbNull
    End Function

    Function GetWeek(D As Date) As Short

        ' bereken week nummer

        GetWeek = D.DayOfYear \ 7 + 1
    End Function

    Public Structure TGameField
        Dim P As Vector3
        Dim N As String ' naam
        Dim T As Short ' top = aantal kaarten op dit veld
    End Structure

    Public SC As Single ' schaal
    Public GS As Single ' raster afstand
    Public Log As String
    Public W(20) As String
    Public V(20) As Double
    Public ParsSkip As Boolean
    Public ParsWord As String
    Public JumpWord As String
    Public ListNaam As String
    Public OX(20) As Single ' delta x
    Public OY(20) As Single ' delta y
    Public OZ(20) As Single ' delta z
    Public AX(20) As Single ' hoek x
    Public AY(20) As Single ' hoek y
    Public AZ(20) As Single ' hoek z
    Public SX(20) As Single ' schaal x
    Public SY(20) As Single ' schaal y
    Public SZ(20) As Single ' schaal z
    Public ParsTo As Boolean
    Public MouseCursor As Cursor
    Public ParsFile(20) As String
    Public FI As Short ' file index
    Public GameField(90) As TGameField
    Public Game As Boolean
    Public WaitKey As Boolean = False
End Module
